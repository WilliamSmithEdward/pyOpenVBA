"""
VBA project parser and compressor.

Implements [MS-OVBA] sufficient for Level 1 read-only extraction and
Level 2 source-replacement write-back.

Stream layout inside the VBA storage (section 2.2)
----------------------------------------------------
  VBA/
    _VBA_PROJECT   - p-code performance cache (opaque; preserve or omit on write)
    dir            - compressed project/module metadata (section 2.3.4.2)
    <stream_name>  - compressed source for each module (section 2.3.4.3)

Critical implementation traps (guide section 31)
-------------------------------------------------
- Module stream name is MODULESTREAMNAME, not the logical module name.
- Source starts at MODULEOFFSET, not byte 0 of the stream.
- Source bytes are MS-OVBA compressed; decompress before decoding.
- Decompressed source bytes are MBCS encoded with the project code page.
- The dir stream itself is also compressed.
- MODULETYPE 0x0022 covers class, document, AND designer modules; use the
  PROJECT stream to distinguish them (not implemented in this version).
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from enum import Enum

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import VBAProjectError

# ---------------------------------------------------------------------------
# Section 2.4 — MS-OVBA compression / decompression
# ---------------------------------------------------------------------------

def copy_token_help(decompressed_current: int, decompressed_chunk_start: int) -> tuple[int, int, int]:
    """
    Return (length_mask, offset_mask, bit_count) for the copy-token at the
    given output-buffer position.  [MS-OVBA] 2.4.1.3.6.

    bit_count = number of bits for the OFFSET field = ceil_log2(difference),
    minimum 4.  As the decoder advances through a 4096-byte decompressed chunk,
    bit_count grows (more back-reference distance available) and the length
    field shrinks accordingly.
    """
    difference = max(1, decompressed_current - decompressed_chunk_start)
    # ceil_log2 via bit_length: (n-1).bit_length() == ceil(log2(n)) for n >= 1
    bit_count = max(4, (difference - 1).bit_length())
    length_mask: int = 0xFFFF >> bit_count
    offset_mask: int = (~length_mask) & 0xFFFF
    return length_mask, offset_mask, bit_count


def decompress(data: bytes) -> bytes:
    """
    Decompress a VBA stream using the MS-OVBA compression algorithm.

    [MS-OVBA] 2.4.1 — CompressedContainer / DecompressedChunk.

    Format:
      Byte 0:    SignatureByte = 0x01
      Bytes 1..: one or more CompressedChunks

    Each CompressedChunk:
      Header u16 LE:
        bits  0-11: CompressedChunkSize = byte_count_of_chunk_data - 1
        bits 12-14: CompressedChunkSignature = 0b011
        bit     15: CompressedChunkFlag (1=compressed, 0=raw)
      Data: (CompressedChunkSize + 1) bytes

    Raw chunk:   Data is exactly 4096 literal bytes.
    Token chunk: Data is groups of (flag_byte, 0-8 tokens).
    """
    if not data or data[0] != 0x01:
        raise VBAProjectError("Invalid compressed stream: missing 0x01 signature byte.")

    pos = 1
    out = bytearray()

    while pos < len(data):
        if pos + 2 > len(data):
            raise VBAProjectError("Truncated compressed stream: missing chunk header.")

        header = int(struct.unpack_from("<H", data, pos)[0])
        chunk_data_size = (header & 0x0FFF) + 1   # data bytes after the header
        chunk_signature = (header >> 12) & 0x7
        chunk_flag = (header >> 15) & 0x1
        pos += 2

        if chunk_signature != 0b011:
            raise VBAProjectError(
                f"Bad compressed chunk signature: expected 0b011, got {chunk_signature:#05b}."
            )

        chunk_end = pos + chunk_data_size
        if chunk_end > len(data):
            raise VBAProjectError(
                f"Truncated chunk: header announces {chunk_data_size} bytes but only "
                f"{len(data) - pos} remain."
            )
        decompressed_chunk_start = len(out)

        if chunk_flag == 0:
            # Raw (uncompressed) chunk — exactly 4096 literal bytes.
            if chunk_data_size != 4096:
                raise VBAProjectError(
                    f"Raw chunk must have exactly 4096 data bytes; got {chunk_data_size}."
                )
            if pos + 4096 > len(data):
                raise VBAProjectError("Truncated raw chunk.")
            out.extend(data[pos: pos + 4096])
            pos += 4096
        else:
            # Token-compressed chunk.
            while pos < chunk_end:
                if pos >= len(data):
                    break
                flag_byte = int(data[pos])
                pos += 1

                for bit in range(8):
                    if pos >= chunk_end or pos >= len(data):
                        break

                    if (flag_byte >> bit) & 1:
                        # Copy token — back-reference into already-decompressed output.
                        if pos + 2 > len(data):
                            raise VBAProjectError("Truncated copy token.")
                        token = int(struct.unpack_from("<H", data, pos)[0])
                        pos += 2

                        length_mask, offset_mask, bit_count = copy_token_help(
                            len(out), decompressed_chunk_start
                        )
                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bit_count)) + 1

                        copy_src = len(out) - offset
                        if copy_src < 0:
                            raise VBAProjectError(
                                "Copy token references before start of output."
                            )
                        if copy_src < decompressed_chunk_start:
                            raise VBAProjectError(
                                "Copy token references before the start of the current chunk."
                            )

                        # Byte-by-byte copy; overlap is intentional and required by spec.
                        for _ in range(length):
                            out.append(out[copy_src])
                            copy_src += 1
                    else:
                        # Literal token.
                        out.append(int(data[pos]))
                        pos += 1

    return bytes(out)


def compress(data: bytes) -> bytes:
    """
    Compress data using the MS-OVBA compression algorithm.

    [MS-OVBA] 2.4.1 — write path.

    Full 4096-byte chunks are stored as raw chunks.
    Partial final chunks are token-compressed via greedy LZ.
    """
    out = bytearray([0x01])   # SignatureByte
    cursor = 0

    while cursor < len(data):
        chunk = data[cursor: cursor + 4096]
        cursor += len(chunk)

        if len(chunk) == 4096:
            # Raw chunk: header flag=0, sig=0b011, size-1=0x0FFF → header=0x3FFF
            out.extend(struct.pack("<H", 0x3FFF))
            out.extend(chunk)
        else:
            encoded = _encode_token_chunk(chunk)
            if len(encoded) > 4096:
                raise VBAProjectError(
                    f"Token-compressed chunk is {len(encoded)} bytes (max 4096)."
                )
            # header: flag=1, sig=0b011, size-1 = len(encoded)-1
            # 0x8000 | 0x3000 = 0xB000
            header = 0xB000 | (len(encoded) - 1)
            out.extend(struct.pack("<H", header))
            out.extend(encoded)

    return bytes(out)


def _encode_token_chunk(chunk: bytes) -> bytes:
    n = len(chunk)
    # Literal-only fits in 4096 bytes when n + ceil(n/8) <= 4096 (i.e. n <= 3640).
    if n + ((n + 7) // 8) <= 4096:
        return _encode_literal_only(chunk)
    return _encode_lz(chunk)


def _encode_literal_only(chunk: bytes) -> bytes:
    out = bytearray()
    i = 0
    while i < len(chunk):
        group = chunk[i: i + 8]
        out.append(0x00)    # all-literal flag byte
        out.extend(group)
        i += len(group)
    return bytes(out)


def _encode_lz(chunk: bytes) -> bytes:
    """Greedy LZ encoder for token-compressed chunks."""
    out = bytearray()
    pos = 0
    chunk_len = len(chunk)

    while pos < chunk_len:
        flag_bits = 0
        tokens: list[bytes] = []

        for bit in range(8):
            if pos >= chunk_len:
                break

            length_mask, offset_mask, bit_count = copy_token_help(pos, 0)
            max_length = length_mask + 3
            max_offset = (offset_mask >> (16 - bit_count)) + 1
            start = max(0, pos - max_offset)

            best_len = 0
            best_offset = 0
            for candidate in range(start, pos):
                match_len = 0
                while (pos + match_len < chunk_len
                       and chunk[candidate + match_len] == chunk[pos + match_len]
                       and match_len < max_length):
                    match_len += 1
                if match_len > best_len:
                    best_len = match_len
                    best_offset = pos - candidate

            if best_len >= 3:
                flag_bits |= (1 << bit)
                offset_bits = ((best_offset - 1) << (16 - bit_count)) & offset_mask
                length_bits = (best_len - 3) & length_mask
                tokens.append(struct.pack("<H", offset_bits | length_bits))
                pos += best_len
            else:
                tokens.append(bytes([chunk[pos]]))
                pos += 1

        out.append(flag_bits)
        for tok in tokens:
            out.extend(tok)

    return bytes(out)


# ---------------------------------------------------------------------------
# Section 2.3.4.2 — dir stream parsing
# ---------------------------------------------------------------------------

class VBAModuleKind(Enum):
    standard = 0x0021   # MODULETYPE_PROCEDURAL
    other = 0x0022      # class, document, or designer — use PROJECT stream to refine


@dataclass
class _ModuleInfo:
    """Internal record accumulated while parsing the dir stream."""
    name: str = ""
    name_unicode: str = ""
    stream_name: str = ""            # MODULESTREAMNAME — the CFB stream to look up
    stream_name_unicode: str = ""
    text_offset: int = 0             # MODULEOFFSET — compressed source starts here
    code_page: int = 1252
    module_kind: VBAModuleKind = VBAModuleKind.standard
    is_read_only: bool = False
    is_private: bool = False


def _parse_dir_stream(raw: bytes) -> tuple[int, list[_ModuleInfo]]:
    """
    Parse a decompressed dir stream.

    Returns (project_code_page, list[_ModuleInfo]).

    Parsing strategy: read records sequentially as (Id:u16, Size:u32, Data:Size bytes).
    Skip unknown records by their Size.  Switch to module parsing at PROJECTMODULES
    marker (0x000F).

    Key record IDs ([MS-OVBA] 2.3.4.2):
      PROJECTCODEPAGE         0x0003  Size=2, CodePage:u16
      PROJECTMODULES          0x000F  Size=2, Count:u16
      PROJECTCOOKIE           0x0013  skip
      MODULENAME              0x0019  MBCS logical module name
      MODULESTREAMNAME        0x001A  MBCS stream name in CFB /VBA/<stream_name>
        reserved partner      0x0032  UTF-16 stream name
      MODULEDOCSTRING         0x001C  skip (MBCS)
        reserved partner      0x0048  skip (UTF-16)
      MODULEOFFSET            0x0031  Size=4, TextOffset:u32
      MODULEHELPCONTEXT       0x001E  skip
      MODULECOOKIE            0x002C  skip
      MODULETYPE standard     0x0021  Size=0
      MODULETYPE other        0x0022  Size=0
      MODULEREADONLY          0x0025  Size=0
      MODULEPRIVATE           0x0028  Size=0
      Module terminator       0x002B  Size=0 (followed by reserved u32=0)
      MODULENAMEUNICODE       0x0047  UTF-16 logical module name
      dir terminator          0x0010
    """
    pos = 0
    code_page = 1252

    def _read_u16() -> int:
        nonlocal pos
        if pos + 2 > len(raw):
            raise VBAProjectError("dir stream truncated while reading u16.")
        v = int(struct.unpack_from("<H", raw, pos)[0])
        pos += 2
        return v

    def _read_u32() -> int:
        nonlocal pos
        if pos + 4 > len(raw):
            raise VBAProjectError("dir stream truncated while reading u32.")
        v = int(struct.unpack_from("<I", raw, pos)[0])
        pos += 4
        return v

    def _skip(n: int) -> bytes:
        nonlocal pos
        chunk = raw[pos: pos + n]
        pos += n
        return chunk

    # ------------------------------------------------------------------
    # Phase 1 — PROJECTINFORMATION + PROJECTREFERENCES
    #
    # Records in this section have heterogeneous layouts: some are
    # (Id, Size, Data) but others (PROJECTDOCSTRING, PROJECTVERSION,
    # PROJECTCONSTANTS, REFERENCENAME, REFERENCECONTROL) carry inline
    # sub-fields that fall outside the Size field.  Fully decoding every
    # variant is large and brittle, and we don't currently need the data.
    #
    # Pragmatic approach: locate PROJECTCODEPAGE (used for MBCS decoding)
    # and PROJECTMODULES (start of the modules section) by signature scan.
    # The PROJECTMODULES record always begins with bytes "0F 00 02 00 00 00"
    # immediately followed by the module-count u16 and then the
    # PROJECTCOOKIE record "13 00 02 00 00 00", so the 14-byte combined
    # signature is highly distinctive.
    # ------------------------------------------------------------------
    # PROJECTCODEPAGE = Id 0x0003, Size=2.
    cp_marker = b"\x03\x00\x02\x00\x00\x00"
    cp_idx = raw.find(cp_marker)
    if cp_idx != -1 and cp_idx + 8 <= len(raw):
        code_page = int(struct.unpack_from("<H", raw, cp_idx + 6)[0])

    # PROJECTMODULES = Id 0x000F, Size=2.
    pm_marker = b"\x0F\x00\x02\x00\x00\x00"
    pm_idx = raw.find(pm_marker)
    if pm_idx == -1:
        raise VBAProjectError("dir stream contains no PROJECTMODULES record.")
    pos = pm_idx + 8   # skip Id(2) + Size(4) + ModuleCount(2)

    # ------------------------------------------------------------------
    # Phase 3 — MODULE records
    # ------------------------------------------------------------------
    modules: list[_ModuleInfo] = []
    current: _ModuleInfo | None = None

    try:
        encoding = f"cp{code_page}"
        "".encode(encoding)
    except LookupError:
        encoding = "latin-1"

    while pos + 2 <= len(raw):
        record_id = _read_u16()

        if record_id == 0x0010:   # dir terminator
            break

        if record_id == 0x0019:   # MODULENAME — starts a new module
            if current is not None:
                modules.append(current)
            current = _ModuleInfo(code_page=code_page)
            record_size = _read_u32()
            current.name = _skip(record_size).decode(encoding, errors="replace")
            continue

        if pos + 4 > len(raw):
            break
        record_size = _read_u32()

        if record_id == 0x0047:   # MODULENAMEUNICODE
            s = _skip(record_size).decode("utf-16-le", errors="replace")
            if current is not None:
                current.name_unicode = s

        elif record_id == 0x001A: # MODULESTREAMNAME (MBCS)
            s = _skip(record_size).decode(encoding, errors="replace")
            if current is not None:
                current.stream_name = s

        elif record_id == 0x0032: # MODULESTREAMNAME reserved (UTF-16)
            s = _skip(record_size).decode("utf-16-le", errors="replace")
            if current is not None:
                current.stream_name_unicode = s

        elif record_id == 0x0031: # MODULEOFFSET
            if record_size == 4 and current is not None:
                current.text_offset = int(struct.unpack_from("<I", raw, pos)[0])
            _skip(record_size)

        elif record_id == 0x0021: # MODULETYPE standard
            if current is not None:
                current.module_kind = VBAModuleKind.standard
            _skip(record_size)

        elif record_id == 0x0022: # MODULETYPE other (class/document/designer)
            if current is not None:
                current.module_kind = VBAModuleKind.other
            _skip(record_size)

        elif record_id == 0x0025: # MODULEREADONLY
            if current is not None:
                current.is_read_only = True
            _skip(record_size)

        elif record_id == 0x0028: # MODULEPRIVATE
            if current is not None:
                current.is_private = True
            _skip(record_size)

        elif record_id == 0x002B: # Module terminator
            _skip(record_size)
            if current is not None:
                modules.append(current)
                current = None

        else:
            _skip(record_size)

    if current is not None and current.name:
        modules.append(current)

    return code_page, modules


# ---------------------------------------------------------------------------
# Public data model
# ---------------------------------------------------------------------------

@dataclass
class VBAModule:
    """A parsed VBA module with its source code."""
    name: str
    stream_name: str
    source: str
    kind: VBAModuleKind = VBAModuleKind.standard
    text_offset: int = 0
    is_read_only: bool = False
    is_private: bool = False
    # Original bytes 0..text_offset of the module stream (performance cache
    # / version-dependent prefix).  Preserved across write-back so that
    # Office's cache invalidation logic operates the same way as it would
    # for an untouched stream.
    prefix_bytes: bytes = field(default=b"", repr=False)
    # Whether the source has been edited and needs to be recompressed on save.
    dirty: bool = field(default=False, repr=False)


@dataclass
class VBAProject:
    """Represents a parsed VBA project."""
    modules: list[VBAModule] = field(default_factory=lambda: [])
    code_page: int = 1252

    def get_module(self, name: str) -> VBAModule:
        needle = name.casefold()
        for m in self.modules:
            if m.name.casefold() == needle:
                return m
        raise KeyError(f"Module not found: {name!r}")

    def module_names(self) -> list[str]:
        return [m.name for m in self.modules]

    # ------------------------------------------------------------------
    # Validation (Gate 19)
    # ------------------------------------------------------------------

    def validate(self, cfb: "CFB | None" = None) -> list[str]:
        """
        Return a list of cross-structure inconsistency messages.

        An empty list means the in-memory project is internally consistent.
        Pass the originating CFB to additionally validate that every
        ``MODULESTREAMNAME`` resolves to a real stream inside ``VBA/``.
        """
        problems: list[str] = []
        seen_names: set[str] = set()
        seen_streams: set[str] = set()
        for m in self.modules:
            key = m.name.casefold()
            if key in seen_names:
                problems.append(f"duplicate module name: {m.name!r}")
            seen_names.add(key)
            skey = m.stream_name.casefold()
            if skey in seen_streams:
                problems.append(f"duplicate module stream name: {m.stream_name!r}")
            seen_streams.add(skey)
            if not m.stream_name:
                problems.append(f"module {m.name!r} has empty MODULESTREAMNAME")
        if cfb is not None:
            try:
                vba_streams = {
                    n.casefold() for n in cfb.list_streams_in_storage("VBA")
                }
            except KeyError:
                problems.append("CFB has no VBA storage")
                return problems
            for m in self.modules:
                if m.stream_name.casefold() not in vba_streams:
                    problems.append(
                        f"module {m.name!r} references missing stream "
                        f"VBA/{m.stream_name!r}"
                    )
        return problems


# ---------------------------------------------------------------------------
# Write-back helpers
# ---------------------------------------------------------------------------

def _encoding_for_codepage(code_page: int) -> str:
    try:
        encoding = f"cp{code_page}"
        "".encode(encoding)
    except LookupError:
        encoding = "latin-1"
    return encoding


def rebuild_module_stream(module: VBAModule, code_page: int) -> bytes:
    """
    Rebuild a module stream by preserving the original ``[0:text_offset]``
    performance-cache prefix and replacing ``[text_offset:]`` with a freshly
    compressed copy of ``module.source``.

    The replacement is byte-exact in the prefix region, so any cache-
    invalidation logic that Office performs on the prefix remains valid.
    """
    encoding = _encoding_for_codepage(code_page)
    source_bytes = module.source.encode(encoding, errors="replace")
    compressed = compress(source_bytes)
    return module.prefix_bytes + compressed


def write_back_modules(cfb: CFB, project: VBAProject) -> None:
    """
    Push every dirty module's source back into ``cfb`` via
    :meth:`CFB.write_stream_in_storage`.  Clean modules are left untouched.
    """
    for m in project.modules:
        if not m.dirty:
            continue
        new_stream = rebuild_module_stream(m, project.code_page)
        cfb.write_stream_in_storage("VBA", m.stream_name, new_stream)
        m.dirty = False


# ---------------------------------------------------------------------------
# Public factory
# ---------------------------------------------------------------------------

def parse_vba_project(cfb: CFB) -> VBAProject:
    """
    Extract and decompress all VBA module sources from a parsed CFB.

    The CFB must be the vbaProject.bin from an xlsm/xlsb, or the whole
    file for an xls workbook.
    """
    try:
        dir_compressed = cfb.get_stream_in_storage("VBA", "dir")
    except KeyError:
        try:
            dir_compressed = cfb.get_stream("dir")
        except KeyError as exc:
            raise VBAProjectError(
                "No 'dir' stream found; not a valid VBA project."
            ) from exc

    dir_raw = decompress(dir_compressed)
    code_page, module_infos = _parse_dir_stream(dir_raw)
    encoding = _encoding_for_codepage(code_page)

    modules: list[VBAModule] = []
    for info in module_infos:
        # CRITICAL: look up by MODULESTREAMNAME, not the logical module name.
        stream_name = info.stream_name or info.name
        try:
            stream_compressed = cfb.get_stream_in_storage("VBA", stream_name)
        except KeyError:
            try:
                stream_compressed = cfb.get_stream(stream_name)
            except KeyError:
                continue

        if info.text_offset > len(stream_compressed):
            raise VBAProjectError(
                f"MODULEOFFSET {info.text_offset} exceeds stream length "
                f"{len(stream_compressed)} for module {info.name!r}."
            )

        compressed_source = stream_compressed[info.text_offset:]
        source_bytes = decompress(compressed_source)
        source = source_bytes.decode(encoding, errors="replace")

        modules.append(VBAModule(
            name=info.name,
            stream_name=stream_name,
            source=source,
            kind=info.module_kind,
            text_offset=info.text_offset,
            is_read_only=info.is_read_only,
            is_private=info.is_private,
            prefix_bytes=stream_compressed[: info.text_offset],
        ))

    return VBAProject(modules=modules, code_page=code_page)
