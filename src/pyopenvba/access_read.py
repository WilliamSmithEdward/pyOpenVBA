"""
pyopenvba.access_read -- Pure-Python read-only access for Microsoft Access (.accdb / .mdb) databases.

Status
------
READ-ONLY. This module exposes a read-only view of Access VBA storage:

    * Read 4 KiB ACE page format (Access 2007+ / Jet 4).
    * Find the VBA project and its modules through ``MSysAccessStorage``,
      as Access does, using the storage engine in :mod:`pyopenvba.access`.
    * Decompress module streams to plain VBA source via
      :func:`pyopenvba.vba.decompress`, in the project's code page.
    * Disassemble Access's flavour of VBA p-code (see :mod:`pyopenvba.vba_pcode`).

Writing Access VBA is :class:`pyopenvba.access.AccessDatabase`'s job.

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
:mod:`pyopenvba.vba`) -- in its own ``MSysAccessStorage`` row, which the
project's dir stream names.  A value of 64 bytes or fewer sits inside the
row; a longer one is laid out across one or more LVAL data pages chained by
a next-page pointer in each page header.

On the **starting** page of a stream, the OVBA signature byte ``0x01`` is
immediately followed by chunk headers and the bytes ``"Attribute VB_Name = ...
"``. On every **continuation** page, the page header bytes 14-15 hold an
offset (little-endian) at which a 4-byte record prefix is followed by the
resumption of the OVBA byte stream. The blob continues to the end of each
chained page.

There is also a secondary plaintext **comment-row index** in some databases
(0xE3 0x00 0x00 0x00 markers + u16 length + ASCII text). This is an Access
find/replace index, NOT the source of truth; callers should ignore it for
source extraction and use :meth:`AccessReader.iter_vba_modules` instead.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from pyopenvba.exceptions import PyOpenVBAError, UnsupportedFormatError
from pyopenvba.vba import VBAReference, encode_mbcs, encoding_for_codepage
from pyopenvba.vba import decompress as _ovba_decompress
from pyopenvba.vba_pcode import (
    DisassembledModule,
    disassemble_module_stream,
)

if TYPE_CHECKING:
    # The storage engine imports this module, so it is imported where used.
    from pyopenvba.access._vba import ModuleStream
    from pyopenvba.access.database import AccessDatabase

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
# reports, macros, modules, relationships, etc.  The storage engine in
# `pyopenvba.access` reads it (see :meth:`AccessReader.iter_msys_objects`).

# MSysObjects ``Type`` values (signed i16). Positive types are
# system-defined container objects; negative types (high bit set)
# tag user content.  Each is what Access writes for the object.
MSYS_TYPE_FORM = -32768          # 0x8000
MSYS_TYPE_MACRO = -32766         # 0x8002
MSYS_TYPE_REPORT = -32764        # 0x8004
MSYS_TYPE_MODULE = -32761        # 0x8007  -- VBA CodeModule
MSYS_TYPE_CONTAINER = 3          # e.g. "Modules", "Forms", "Reports" hubs
MSYS_TYPE_TABLE = 1
MSYS_TYPE_QUERY = 5
MSYS_TYPE_DATABASE = 8

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


class AccessReader:
    """
    Read-only entry point for an Access database file. Construction parses the
    file header and validates the ACE/Jet signature. Higher-level methods
    (``module_names``, source extraction, etc.) are implemented incrementally.

    Usage::

        with AccessReader("database.accdb") as db:
            print(db.format)            # "ace" or "jet4"
            print(db.page_count)
    """

    @classmethod
    def create_new(cls, path: str | Path) -> AccessReader:
        """Create a blank ``.accdb`` at ``path`` and return a reader for it.

        The bytes come from a template captured from a database Access
        authored itself, so the result opens cleanly. It holds one
        standard module, ``Module1``, containing an empty ``Main``
        function -- a starting point that already has the VBA project,
        module and procedure structure a database needs, none of which
        can be synthesised from nothing.

        Mirrors :meth:`pyopenvba.ExcelFile.create_new` and its Word and
        PowerPoint counterparts. ``path`` is overwritten if it exists.
        """
        from pyopenvba._templates import EMPTY_ACCDB_BYTES

        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_ACCDB_BYTES)
        return cls(target)

    def __init__(self, path: str | Path) -> None:
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

    def iter_source_rows(self, code_page: int = 1252) -> Iterator[SourceRow]:
        """
        Yield every VBA source-line row found anywhere in the database, in
        file-offset order.

        Each row carries a row-type marker (currently only the comment
        marker has been observed and decoded), a 16-bit text length, and an
        MBCS payload. The leading "' " of stored comment lines is *not*
        included in ``text`` -- callers should prepend it when reconstructing
        the line as it would appear in the VBA editor.

        ``code_page`` is the project's ``PROJECTCODEPAGE`` (see
        :meth:`read_project_info`); it is recorded on each row so the text
        decodes correctly. It defaults to 1252, which is right for
        Western-European projects and harmless for pure-ASCII source.

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
            # Reject obviously non-source payloads. Source text is stored
            # MBCS-encoded in the project's code page, so bytes above 0x7E
            # are legitimate (accented Latin, Cyrillic, CJK); only C0/C1
            # controls other than tab, CR and LF disqualify a payload.
            if any(
                b < 0x09 or 0x0E <= b <= 0x1F or b == 0x7F
                or (0x0A < b < 0x0D)
                for b in payload
            ):
                continue
            yield SourceRow(
                offset=j,
                row_type="comment",
                length=length,
                text=payload,
                code_page=code_page,
            )

    def __enter__(self) -> AccessReader:
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        # Nothing to release: we read the file fully into memory at __init__.
        return None

    # ------------------------------------------------------------------
    # The VBA project, through MSysAccessStorage.
    # ------------------------------------------------------------------
    #
    # Access keeps the project as one `MSysAccessStorage` row per stream:
    # the standard MS-OVBA "dir" stream (section 2.3.4.2 of MS-OVBA),
    # which parses with `pyopenvba.vba`, and a row per module that the dir
    # stream names.  The storage engine in `pyopenvba.access` reads that
    # table, so the project is found where Access looks for it.  A scan of
    # the long-value pages is no substitute: they keep copies Access has
    # let go of, a short value sits inside its row instead, and a long one
    # is a chain whose head starts with a pointer (GitHub issue #33).

    def _database(self) -> AccessDatabase | None:
        """The storage engine over this file's bytes, or ``None`` when the
        file holds no VBA project (see
        :meth:`pyopenvba.access.AccessDatabase.has_vba_project`)."""
        from pyopenvba.access.database import AccessDatabase  # it imports this module

        database = AccessDatabase(bytes(self._data))
        return database if database.has_vba_project() else None

    def read_project_info(self) -> AccessVBAProject:
        """Parse and return the project-level VBA metadata embedded in
        this database.

        Raises :class:`AccessError` if the database holds no VBA project.
        """
        from pyopenvba.vba import parse_dir_stream

        database = self._database()
        if database is None:
            raise AccessError(f"{self.path.name!r} holds no VBA project")
        raw, home = database.dir_stream()
        info, mods = parse_dir_stream(raw)
        return AccessVBAProject(
            catalog_page=home.page,
            catalog_slot=home.slot,
            catalog_raw_size=len(raw),
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

    def _project_streams(self) -> list[tuple[str, bytes, int, int]]:
        """The project's streams as ``(name, value, page, slot)``, ordered
        by where each value starts; empty when there is no VBA project.
        See :meth:`pyopenvba.access.AccessDatabase.project_streams`."""
        database = self._database()
        if database is None:
            return []
        streams = [
            (name, value, home.page, home.slot)
            for name, value, home in database.project_streams()
        ]
        return sorted(streams, key=lambda stream: (stream[2], stream[3]))

    # ------------------------------------------------------------------
    # Phase 4 RE: authoritative VBA p-code stream
    # ------------------------------------------------------------------
    #
    # A database whose VBA Access has compiled carries `__SRP_*` streams
    # whose bytes begin with the magic header `72 55 40 ...` ("rU@"), and
    # each compiled module has one marked module-active.  That row is the
    # authoritative compiled-bytecode store -- decompiling and
    # re-displaying VBA source in the Access editor reads from it, not
    # from the OVBA cache. The OVBA cache (the stream `iter_vba_modules`
    # reads) is a passive plaintext mirror Access keeps for
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

    # The p-code header is 12 bytes. Every rU@-headed stream starts
    # with the 4-byte signature ``72 55 40 00`` followed by 8 more
    # bytes whose structure encodes the stream's role:
    #
    #   bytes  0..3   : signature 'rU@\x00'
    #   bytes  4..7   : reserved / zero in the corpus
    #   bytes  8..15  : u64 (LE) -- 0x4000 for the *module-active*
    #                   bytecode row, 0 for every other rU@ row
    #
    # The 0x4000 at offset 10 is the deterministic structural marker
    # that distinguishes a row Access actually executes from the stubs
    # and project/system bootstrap rows kept alongside it.  Verified
    # across the 15-sample corpus (samples 010-051), and on databases
    # Access wrote with two and with thirteen compiled modules, which
    # carry that many.
    _PCODE_MAGIC = b"\x72\x55\x40\x00"   # 'rU@\x00'
    _PCODE_ACTIVE_PREFIX = (
        b"\x72\x55\x40\x00\x00\x00\x00\x00\x00\x00\x40\x00"
    )

    def iter_pcode_streams(self) -> tuple[AccessVBAPCodeStream, ...]:
        """Return every ``rU@``-headed VBA p-code stream the project holds.

        Each compiled module has one *module-active* stream, the
        bytecode Access executes; the others are stubs or project/system
        bootstrap streams. :meth:`read_module_pcode_stream` fetches the
        active one of a project with a single compiled module.

        Raises :class:`AccessError` if the project holds none, as a
        project written only as source does until Access compiles it.
        """
        hits = [
            AccessVBAPCodeStream(page=page, slot=slot, raw=value)
            for _name, value, page, slot in self._project_streams()
            if value.startswith(self._PCODE_MAGIC)
        ]
        if not hits:
            raise AccessError(
                f"no VBA p-code streams (header 'rU@') found in "
                f"{self.path.name!r}; its VBA may never have been compiled"
            )
        return tuple(hits)

    def read_module_pcode_stream(self) -> AccessVBAPCodeStream:
        """Return the *module-active* VBA p-code stream, identified by the
        structural 12-byte prefix ``72 55 40 00 00 00 00 00 00 00 40
        00`` (byte at offset 10 is ``0x40`` rather than ``0x00``).

        This is the deterministic discriminator that separates the
        active compiled bytecode from the stub and bootstrap streams
        Access keeps alongside it.

        Raises :class:`AccessError` if no stream matches the active
        prefix, or if several do, as they do when more than one module
        is compiled.
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
    # Each string literal is stored as ``0B <u32 LE byte-count>
    # <UTF-16-LE bytes>``.  The leading ``0B`` tag distinguishes literal
    # records from the other entries in the same stream.
    _STRING_LITERAL_TAG = 0x0B

    def find_interned_strings(self) -> tuple[AccessVBAInternedString, ...]:
        """Scan the project's streams for VBA string-literal records of
        the form ``0B <u32 LE byte-count> <UTF-16-LE bytes>``.

        Access writes them into the ``__SRP_0`` stream it keeps beside a
        compiled project's ``rU@`` execodes, one for each literal the
        compiled code holds, a doubled quote made one, among entries of
        other kinds.  The empty literal is a record of zero bytes, which
        this scan does not report.  A project saved without its
        ``__SRP_*`` streams has no such table, and its literals are only
        in each module's own p-code, as ``LitStr`` instructions (see
        :meth:`disassemble_module`).

        This is a deterministic content-based scan -- no offsets are
        hard-coded.  The decoder walks every stream the project holds,
        each whole, and yields each valid literal record it finds.  Only
        streams the project still holds count: the long-value pages also
        keep the literals of code that was deleted, and those are not the
        project's.

        A record is accepted only when:

        * the byte-count is even and non-zero,
        * the byte-count fits in the stream,
        * the decoded UTF-16-LE bytes form a valid Python ``str``,
        * and the decoded string contains no NUL characters
          (filters out structural padding that happens to start with
          ``0B``).

        Returns a tuple of :class:`AccessVBAInternedString` records,
        each carrying where the stream starts, the record's byte offset
        in the stream, and the decoded value.
        """
        out: list[AccessVBAInternedString] = []
        for _name, buf, page, slot in self._project_streams():
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
    # In every Access database we inspected, the stream carrying a
    # module's OVBA-compressed VBA source ALSO contains -- ahead of the
    # source, before its MODULEOFFSET -- the standard Office VBA module
    # stream's "PerformanceCache" region, recognisable by the
    # well-known ``0xCAFE`` magic word. This is the same per-line
    # p-code layout described in [MS-OVBA] section 2.3.4.3 and
    # disassembled by the public `pcodedmp` tool. It is *NOT* the
    # ``rU@``-prefixed bytecode (read by
    # :meth:`read_module_pcode_stream`); the ``rU@`` stream is the
    # Access-specific cached/execodes form. Both forms coexist in the
    # database; Access uses ``rU@`` at runtime, but the canonical VBA7
    # p-code -- portable across all Office hosts -- lives here.

    def find_module_streams(self) -> tuple[AccessVBAModuleStream, ...]:
        """Return the standard Office VBA module-stream bytes for
        every VBA module in the database.

        For each VBA module, the stream containing its OVBA-
        compressed source also contains -- ahead of the source -- the
        standard module stream's binary
        ``PerformanceCache`` region. That region is recognisable by
        the ``0xCAFE`` magic word and contains the canonical VBA7
        p-code (per-line opcodes), in the exact layout defined by
        [MS-OVBA] and consumed by public disassemblers.

        Returns one :class:`AccessVBAModuleStream` per module,
        carrying the source page, slot, raw row bytes, and the
        in-row offset of the ``0xCAFE`` magic. The raw bytes can be
        fed directly to any disassembler that expects an Office VBA
        module stream.

        Each stream is the one the project's dir stream names, and the
        ``0xCAFE`` word is looked for only ahead of its MODULEOFFSET,
        where the compiled region ends.  A module stored as source alone,
        as :class:`pyopenvba.access.AccessDatabase` writes one, has no
        compiled region until Access compiles it, and is left out.
        """
        database = self._database()
        if database is None:
            return ()
        results: list[AccessVBAModuleStream] = []
        for stream in database.module_streams():
            if stream.home is None:
                continue
            cafe = stream.data.find(b"\xfe\xca", 0, stream.offset)
            if cafe < 0:
                continue
            page, slot = stream.home
            results.append(
                AccessVBAModuleStream(
                    page=page, slot=slot, raw=stream.data, cafe_offset=cafe,
                    name=stream.name,
                )
            )
        return tuple(results)

    def identifiers(self) -> tuple[AccessVBAIdentifier, ...]:
        """Enumerate every project-level identifier name decoded from
        the project's ``_VBA_PROJECT`` stream.

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

        Returns an empty tuple if the project holds no ``_VBA_PROJECT``
        stream opening with ``CC 61`` (no VBA project, or a damaged one).

        Note: this is a project-wide *inventory*; the ``name_id``
        u16 operands in p-code do **NOT** index into this table
        directly (they index a per-procedure reference table -- a
        future RE deliverable). The inventory is still useful for
        diagnostic and auditing purposes (e.g. listing every
        intrinsic a project calls).
        """
        stream = next(
            (value for name, value, _page, _slot in self._project_streams() if name == "_VBA_PROJECT"),
            None,
        )
        if stream is None or not stream.startswith(b"\xcc\x61"):
            return ()
        return _parse_vba_project_identifiers(stream, self.read_project_info().code_page)

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
        # Match by name: several modules can share one LVAL page, so a
        # page-keyed lookup would silently return a neighbour's p-code.
        by_name = {s.name: s for s in self.find_module_streams()}
        stream = by_name.get(name)
        if stream is not None:
            return disassemble_module_stream(stream.raw, is_64bit=is_64bit)
        if name in self.vba_module_names():
            raise AccessError(
                f"module {name!r} has no compiled p-code "
                "(no 0xCAFE region in carrier row)"
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
        return {
            stream.name: disassemble_module_stream(
                stream.raw, is_64bit=is_64bit
            )
            for stream in self.find_module_streams()
        }

    def iter_vba_modules(self) -> Iterator[VBAModule]:
        """
        Yield every VBA module in this database, in the order its
        project lists them.

        Each module is found as Access finds it: the project's dir stream
        names the ``MSysAccessStorage`` row holding the module's stream,
        and its MODULEOFFSET says where the compressed source starts.  The
        text is decoded in the project's code page, not as latin-1.  A
        module the dir stream lists whose row is missing is left out, and
        a database with no VBA project yields nothing.
        """
        for _stream, module in self._read_modules():
            yield module

    def _read_modules(self) -> Iterator[tuple[ModuleStream, VBAModule]]:
        """Each module's stream beside the module it reads as."""
        from pyopenvba.access._vba import split_source

        database = self._database()
        if database is None:
            return
        for stream in database.module_streams():
            if stream.home is None:
                continue
            raw = _ovba_decompress(stream.data[stream.offset :], stream_name=stream.stream_name)
            attributes, body = split_source(raw.decode(stream.encoding, errors="replace"))
            yield stream, VBAModule(
                name=stream.name,
                start_offset=stream.home[0] * ACE_PAGE_SIZE,
                raw_blob_size=len(stream.data) - stream.offset,
                decompressed_size=len(raw),
                attributes_text="\r\n".join(attributes),
                source="\r\n".join(body),
            )

    def vba_module_names(self) -> list[str]:
        """
        Return the name of every VBA module in this database, in the
        order its project's dir stream lists them; empty when the
        database holds no VBA project.
        """
        database = self._database()
        if database is None:
            return []
        return [stream.name for stream in database.module_streams()]

    def read_vba_module(self, name: str) -> str:
        """
        Return the user-visible source text of the named module (without
        the leading ``Attribute VB_*`` preamble lines, with ``\\r\\n``
        line endings preserved).

        Raises :class:`AccessError` if no module with that name is found.
        """
        return self._vba_module(name).source

    def _vba_module(self, name: str) -> VBAModule:
        for module in self.iter_vba_modules():
            if module.name == name:
                return module
        raise AccessError(f"VBA module {name!r} not found in {self.path.name!r}")

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


    def read_vba_module_with_attributes(self, name: str) -> str:
        """Like :meth:`read_vba_module` but returns the full module
        text including the leading ``Attribute VB_*`` preamble (and
        the ``VERSION ... CLASS`` block for class modules), separated
        from the body by the canonical CRLF terminator.

        Raises :class:`AccessError` if no module with that name exists.
        """
        m = self._vba_module(name)
        attrs = m.attributes_text
        if attrs and not attrs.endswith("\r\n"):
            attrs = attrs + "\r\n"
        return attrs + m.source


    # ------------------------------------------------------------------
    # Excel-parallel ergonomic API: get/set/vba_modules/push/pull.
    # ------------------------------------------------------------------

    def get_module(self, name: str) -> str:
        """Return the body source of module ``name`` (no attribute
        preamble). Excel-parallel alias for :meth:`read_vba_module`."""
        return self.read_vba_module(name)


    def vba_modules(self) -> dict[str, str]:
        """Return ``{module_name: body_source}`` for every module in
        the catalog. Excel-parallel."""
        return {m.name: m.source for m in self.iter_vba_modules()}

    def pull_modules(
        self,
        dest_dir: str | Path,
        *,
        encoding: str = "utf-8",
        overwrite: bool = True,
    ) -> list[Path]:
        """Export every VBA module body to ``dest_dir`` as one file per
        module (``.bas`` for std modules, ``.cls`` for class modules).
        Excel-parallel. Returns the list of files written.

        Like Excel's :meth:`pull_modules`, this writes only the user-
        visible *body* (no ``Attribute VB_*`` preamble). To include the
        preamble use :meth:`export_modules` with
        ``include_attributes=True``.

        It is :meth:`pyopenvba.access.AccessDatabase.pull_modules`, so the
        two write the same files.
        """
        database = self._database()
        if database is None:
            Path(dest_dir).mkdir(parents=True, exist_ok=True)
            return []
        return database.pull_modules(dest_dir, encoding=encoding, overwrite=overwrite)


    # ------------------------------------------------------------------

    def export_module(self, name: str) -> str:
        """
        Return the user-visible source text of a single module by name.

        Identical to :meth:`read_vba_module`; provided for symmetry with
        :meth:`export_modules` and :meth:`import_module`.
        """
        return self.read_vba_module(name)

    def export_modules(
        self,
        dest_dir: str | Path,
        *,
        include_attributes: bool = False,
    ) -> list[Path]:
        """
        Write every module to ``dest_dir`` as one file per module.

        Class modules are written as ``<name>.cls``; everything else as
        ``<name>.bas``. The leading ``Attribute VB_*`` preamble is omitted
        by default (this matches what the VBA editor shows on screen); set
        ``include_attributes=True`` to round-trip the raw stream.  The
        text is written in the project's code page, the bytes the stream
        holds.

        Returns the list of files written. The destination directory is
        created if it does not exist.
        """
        out_dir = Path(dest_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        written: list[Path] = []
        for stream, module in self._read_modules():
            ext = ".cls" if stream.kind == "class" else ".bas"
            target = out_dir / (module.name + ext)
            if include_attributes:
                attrs = module.attributes_text
                if attrs and not attrs.endswith("\r\n"):
                    attrs += "\r\n"
                body = attrs + module.source
            else:
                body = module.source
            target.write_bytes(encode_mbcs(body, stream.encoding))
            written.append(target)
        return written


    # ------------------------------------------------------------------
    # MSysObjects (Jet/ACE system catalog) -- read path
    # ------------------------------------------------------------------

    def iter_msys_objects(self) -> Iterator[AccessSysObject]:
        """Iterate every persistent object listed in the .accdb's
        ``MSysObjects`` system catalog.

        Yields one :class:`AccessSysObject` per live row, read by the
        storage engine, a row moved to an overflow page included. Each
        row identifies a Table, Query, Form, Report, Macro, VBA Module,
        or system Container.  A hand decoder here once skipped moved
        rows, cut some names short and dropped others it could not
        bound.

        Use :meth:`find_msys_module` for the common case of locating
        a single VBA code module by name. Use :meth:`iter_msys_modules`
        to enumerate only VBA module rows.
        """
        from pyopenvba.access.database import AccessDatabase  # it imports this module

        for entry in AccessDatabase(bytes(self._data)).catalog():
            yield AccessSysObject(
                id_=entry.id & 0xFFFFFFFF,
                parent_id=entry.parent_id & 0xFFFFFFFF,
                type_=entry.type,
                flags=entry.flags,
                name=entry.name,
                page=entry.page,
                slot=entry.row,
            )

    def msys_objects(self) -> tuple[AccessSysObject, ...]:
        """Return all MSysObjects rows as a tuple (materialised list of
        :meth:`iter_msys_objects`)."""
        return tuple(self.iter_msys_objects())

    def iter_msys_modules(self) -> Iterator[AccessSysObject]:
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
    ) -> AccessSysObject | None:
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

    def find_msys_module(self, name: str) -> AccessSysObject | None:
        """Return the MSysObjects row for the VBA code module called
        ``name`` (case-insensitive), or ``None`` if no such module
        exists in the system catalog."""
        return self.find_msys_object(name, type_=MSYS_TYPE_MODULE)


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
        page: ACE 4 KiB page number of the row's home slot, where index
            entries name it; a row moved to an overflow page keeps it.
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
    code_page: int = 1252

    def to_source_line(self) -> str:
        """Reconstruct the source line as it would appear in the VBA editor.

        ``text`` is raw MBCS bytes in the project's code page, so it is
        decoded with that page rather than ASCII; decoding as ASCII raised
        ``UnicodeDecodeError`` on any accented or non-Latin comment.
        """
        encoding = encoding_for_codepage(self.code_page)
        text = self.text.decode(encoding, errors="replace")
        return "'" + text if self.row_type == "comment" else text


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
    is loaded separately via :meth:`AccessReader.iter_vba_modules`.

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
    dir stream is stored OVBA-compressed in its ``MSysAccessStorage`` row;
    ``catalog_page`` / ``catalog_slot`` say where that value starts, its
    first LVAL row, for diagnostic purposes.
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
    together with where the stream holding them starts.

    The first four bytes are always ``72 55 40 00`` ('rU@\\x00'). The
    full opcode field guide is being reverse-engineered; see
    ``docs/access_pcode_re.md``.

    Attributes:
        page: ACE 4 KiB page number of the LVAL row where the stream
            starts.
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
    from the storage row that also carries its OVBA-compressed source.

    Recognisable by the ``0xCAFE`` magic word in ``raw[cafe_offset:]``
    that marks the start of the per-line p-code region (see [MS-OVBA]
    section 2.3.4.3). The full byte layout matches what public VBA
    disassemblers (e.g. ``pcodedmp``) consume.

    This is the *canonical* portable VBA p-code -- the form that any
    Office host running VBA7 can execute. It coexists with the
    Access-specific ``rU@``-prefixed cached form (see
    :class:`AccessVBAPCodeStream`).

    Attributes:
        page: ACE 4 KiB page number containing the LVAL row where the
            stream starts.
        slot: Slot index within ``page``. Several modules commonly share
            a page, so ``page`` alone does not identify a module.
        raw: The module stream, whole, however many LVAL rows it spans.
            The module-stream-format region runs from offset 0 through
            the start of the OVBA compressed source.
        cafe_offset: In-row byte offset of the ``0xCAFE`` magic word
            that opens the p-code region.
        name: Module name, from the project's dir stream.
    """

    page: int
    slot: int
    raw: bytes
    cafe_offset: int
    name: str = ""


@dataclass(frozen=True)
class AccessVBAInternedString:
    """A single VBA string-literal record decoded from the project's
    intern table.

    Each literal is stored as a ``0B <u32 LE byte-count> <UTF-16-LE>``
    record in the ``__SRP_0`` stream of a compiled project (see
    :meth:`AccessReader.find_interned_strings`).  The ``rU@`` execodes
    refer to a literal by its slot; its text is here, and in its
    module's own p-code as a ``LitStr`` instruction.

    Attributes:
        page: ACE page number of the LVAL row where the stream carrying
            the record starts.
        slot: Slot index within ``page``.
        offset: Byte offset of the ``0B`` tag within the stream.
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

    The Access ``_VBA_PROJECT`` payload is stored uncompressed, and its
    first two bytes are the magic ``CC 61``. Near the tail of the
    stream the host emits a list of identifier records, one
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

    Most records are addressed by position: the ``name`` operand of a
    compiled p-code ``Ld`` / ``St`` instruction is ``524 + 2*index``
    (measured across Access-built databases, 2026-08). A few names bind
    instead to a pre-existing low-numbered slot and are stored in a
    variant record that carries that slot explicitly; those set
    :attr:`slot`, are excluded from the positional numbering, and are
    addressed as ``2*slot + 2``.

    Attributes:
        index: 0-based position within the identifier table, or ``-1``
            for a record that carries its own :attr:`slot` and so takes
            no position.
        type_byte: The single type byte preceding the name in the
            record.
        name: ASCII name, decoded from the on-disk byte payload.
        id_low: 16-bit ID cookie that follows the name on disk. Zero
            for slotted records, which have no such trailer.
        prefix: Any extra descriptor bytes seen before this record that
            could not be parsed as another ``<len><type><name>`` entry.
            Empty for fully canonical records.
        slot: Explicit operand slot for records that carry one, else
            ``None`` for the usual positional records.
    """

    index: int
    type_byte: int
    name: str
    id_low: int
    prefix: bytes
    slot: int | None = None


def _parse_vba_project_identifiers(
    stream: bytes,
    code_page: int = 1252,
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
        # A few identifiers carry their own operand slot instead of
        # being addressed by position:
        #     00 00 <u16 slot> <u8 len> 80 <6B descriptor> <name>
        # There is no trailing id / 0x10 0x00 pair, which is why the
        # canonical walk below rejects the record. Such an entry must
        # NOT advance ``index``: compiled p-code addresses positional
        # records as ``524 + 2*index``, so counting one here would
        # misname every identifier after it.
        if (
            stream[pos] == 0x00
            and stream[pos + 1] == 0x00
            and pos + 12 <= end
            and stream[pos + 5] == 0x80
            and stream[pos + 8:pos + 10] == b"\xff\x03"
        ):
            slot_len = stream[pos + 4]
            slot_name_end = pos + 12 + slot_len
            if (
                0 < slot_len < 64
                and slot_name_end <= end
                and all(
                    stream[i] >= 0x20 and stream[i] != 0x7F
                    for i in range(pos + 12, slot_name_end)
                )
            ):
                out.append(
                    AccessVBAIdentifier(
                        index=-1,
                        type_byte=stream[pos + 5],
                        name=stream[pos + 12:slot_name_end].decode(
                            encoding_for_codepage(code_page),
                            errors="replace",
                        ),
                        id_low=0,
                        prefix=pending_prefix,
                        slot=int.from_bytes(
                            stream[pos + 2:pos + 4], "little"
                        ),
                    )
                )
                pending_prefix = b""
                pos = slot_name_end
                continue
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
                stream[i] >= 0x20 and stream[i] != 0x7F
                for i in range(name_start, name_end)
            )
        ):
            name = stream[name_start:name_end].decode(
                encoding_for_codepage(code_page), errors="replace"
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
