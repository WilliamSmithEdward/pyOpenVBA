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

import codecs
import struct
import unicodedata
import warnings
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from enum import Enum
from typing import TYPE_CHECKING

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import VBAProjectError

if TYPE_CHECKING:
    from pyopenvba.vba_pcode import DisassembledModule

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


def decompress(
    data: bytes, *, stream_name: str = "<unknown>", max_bytes: int | None = None
) -> bytes:
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

    ``stream_name`` is purely contextual: it is embedded in any
    :class:`VBAProjectError` raised here so callers can identify which
    CFB stream a malformed chunk came from.

    ``max_bytes``, when given, stops decompression at the first chunk
    boundary at or beyond that many output bytes and returns the
    chunk-aligned prefix.  Copy tokens never reference across chunk
    boundaries (this decoder enforces it), so the prefix is
    byte-identical to the same range of a full decompression.  Useful
    for reading a module's ``Attribute VB_*`` header without paying for
    the whole stream.
    """
    out, _ = _decompress(data, stream_name=stream_name, max_bytes=max_bytes)
    return out


def _decompress(
    data: bytes, *, stream_name: str, max_bytes: int | None
) -> tuple[bytes, bool]:
    """Core decoder behind :func:`decompress`.

    Returns ``(decompressed, consumed_all)`` where ``consumed_all`` is
    True when the entire compressed input was processed — i.e. the
    returned bytes are the complete decompression, not a prefix.

    Performance shape (measured; see issue #5): output is produced with
    slice operations wherever the spec allows — non-overlapping copy
    tokens move as one slice, runs of literal tokens within a flag byte
    extend once — and the copy-token masks are recomputed only when the
    chunk-local output size crosses a power of two, since
    :func:`copy_token_help` depends on nothing else.  Overlapping
    copies (offset < length) keep the spec's byte-at-a-time semantics,
    which are load-bearing there.
    """

    def _err(msg: str, offset: int) -> VBAProjectError:
        return VBAProjectError(
            f"{msg} [stream={stream_name!r}, offset={offset}]"
        )

    if not data or data[0] != 0x01:
        raise _err("Invalid compressed stream: missing 0x01 signature byte.", 0)

    pos = 1
    n = len(data)
    out = bytearray()

    while pos < n:
        if max_bytes is not None and len(out) >= max_bytes:
            return bytes(out), False

        if pos + 2 > n:
            raise _err("Truncated compressed stream: missing chunk header.", pos)

        header = data[pos] | data[pos + 1] << 8
        chunk_data_size = (header & 0x0FFF) + 1   # data bytes after the header
        chunk_signature = (header >> 12) & 0x7
        chunk_flag = (header >> 15) & 0x1
        header_offset = pos
        pos += 2

        if chunk_signature != 0b011:
            raise _err(
                f"Bad compressed chunk signature: expected 0b011, "
                f"got {chunk_signature:#05b}.",
                header_offset,
            )

        chunk_end = pos + chunk_data_size
        if chunk_end > n:
            raise _err(
                f"Truncated chunk: header announces {chunk_data_size} bytes "
                f"but only {n - pos} remain.",
                header_offset,
            )
        decompressed_chunk_start = len(out)

        if chunk_flag == 0:
            # Raw (uncompressed) chunk — exactly 4096 literal bytes.
            if chunk_data_size != 4096:
                raise _err(
                    f"Raw chunk must have exactly 4096 data bytes; "
                    f"got {chunk_data_size}.",
                    header_offset,
                )
            if pos + 4096 > n:
                raise _err("Truncated raw chunk.", pos)
            out.extend(data[pos: pos + 4096])
            pos += 4096
        else:
            # Token-compressed chunk.  Masks depend only on the
            # chunk-local output offset; recompute at power-of-two
            # crossings instead of per token.
            length_mask, offset_mask, bit_count = copy_token_help(
                len(out), decompressed_chunk_start
            )
            next_threshold = (1 << bit_count) + 1

            while pos < chunk_end:
                if pos >= n:
                    break
                flag_byte = data[pos]
                pos += 1

                bit = 0
                while bit < 8:
                    if pos >= chunk_end or pos >= n:
                        break

                    if (flag_byte >> bit) & 1:
                        # Copy token — back-reference into already-decompressed output.
                        if pos + 2 > n:
                            raise _err("Truncated copy token.", pos)
                        token = data[pos] | data[pos + 1] << 8
                        pos += 2

                        if len(out) - decompressed_chunk_start >= next_threshold:
                            length_mask, offset_mask, bit_count = copy_token_help(
                                len(out), decompressed_chunk_start
                            )
                            next_threshold = (1 << bit_count) + 1
                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bit_count)) + 1

                        copy_src = len(out) - offset
                        if copy_src < 0:
                            raise _err(
                                "Copy token references before start of output.",
                                pos - 2,
                            )
                        if copy_src < decompressed_chunk_start:
                            raise _err(
                                "Copy token references before the start of the "
                                "current chunk.",
                                pos - 2,
                            )

                        if offset >= length:
                            # Source range is fully materialized: one slice.
                            out += out[copy_src: copy_src + length]
                        else:
                            # Overlapping copy; byte-at-a-time is required
                            # by spec semantics (the pattern repeats).
                            for _ in range(length):
                                out.append(out[copy_src])
                                copy_src += 1
                        bit += 1
                    else:
                        # Run of literal tokens: consecutive clear bits in
                        # this flag byte copy verbatim as one slice.
                        run_start = pos
                        while (bit < 8
                               and not ((flag_byte >> bit) & 1)
                               and pos < chunk_end
                               and pos < n):
                            pos += 1
                            bit += 1
                        out += data[run_start:pos]

    return bytes(out), True


def compress(data: bytes) -> bytes:
    """
    Compress data using the MS-OVBA compression algorithm.

    [MS-OVBA] 2.4.1 — write path.

    Every chunk is emitted as token-compressed (CompressedChunkFlag=1).
    Although the spec permits raw (flag=0) chunks of exactly 4096 bytes,
    Office itself never writes them in CompressedSourceCode streams: every
    chunk in an Excel-authored module stream is token-compressed.
    Empirically, the Excel/VBE reader refuses to load module source whose
    first chunk is raw ("An error occurred while loading <Module>"), even
    though the bytes round-trip through a spec-compliant decompressor.
    We therefore always emit token-compressed chunks and only fall back to
    a raw chunk if the LZ encoding overflows 4096 bytes — which does not
    happen for realistic VBA source text but is a valid escape hatch for
    adversarial high-entropy input.
    """
    out = bytearray([0x01])   # SignatureByte
    cursor = 0

    while cursor < len(data):
        chunk = data[cursor: cursor + 4096]
        cursor += len(chunk)

        encoded = _encode_token_chunk(chunk)
        if len(encoded) <= 4096:
            # header: flag=1, sig=0b011, size-1 = len(encoded)-1
            # 0x8000 | 0x3000 = 0xB000
            header = 0xB000 | (len(encoded) - 1)
            out.extend(struct.pack("<H", header))
            out.extend(encoded)
            continue

        # Token encoding overflowed (high-entropy input). Spec permits a
        # raw chunk here, but only at a strict 4096 data bytes; smaller
        # final chunks cannot be raw.
        if len(chunk) != 4096:
            raise VBAProjectError(
                f"Final partial chunk ({len(chunk)} bytes) encoded to "
                f"{len(encoded)} bytes, exceeding the 4096-byte chunk "
                f"limit. The data cannot be represented as a single "
                f"CompressedChunk."
            )
        # Raw chunk: header flag=0, sig=0b011, size-1=0x0FFF → header=0x3FFF
        out.extend(struct.pack("<H", 0x3FFF))
        out.extend(chunk)

    return bytes(out)


def _encode_token_chunk(chunk: bytes) -> bytes:
    # Always run the greedy LZ encoder. Microsoft Access byte-validates
    # the OVBA cache blob and rejects modules whose CompressedChunk bytes
    # don't match what its own compressor would have produced -- even
    # when the decompressed plaintext is identical. A literal-only
    # fast-path (legal per [MS-OVBA] 2.4.1) emits valid-but-different
    # bytes for short chunks (<= 3640 bytes) and trips Access's
    # "Error accessing file. Network connection may have been lost."
    # in the VBE. Empirically the greedy LZ encoder produces output
    # that matches Access's encoding byte-for-byte on samples 010-051.
    # If LZ encoding overflows 4096 bytes (adversarial high-entropy
    # input), fall back to a raw chunk -- this only fires for inputs
    # that cannot be represented as a token-compressed chunk at all.
    encoded = _encode_lz(chunk)
    if len(encoded) <= 4096:
        return encoded
    # Caller will see len > 4096 and emit a raw chunk if the chunk is
    # exactly 4096 bytes. This branch is unreachable for realistic VBA
    # source text.
    return encoded


def _encode_lz(chunk: bytes) -> bytes:
    """Greedy LZ encoder for token-compressed chunks.

    Byte-for-byte equivalent to the naive full-window scan (each
    position considers every earlier position, ascending, keeping the
    first candidate with a strictly greater capped match length -- so
    ties go to the FARTHEST candidate), but replaces the O(window) scan
    with a 3-gram position index.  The reduction is sound because:

    * a copy token requires a match length of at least 3, so any
      winning candidate shares its first three bytes with the target
      and appears in the index bucket for that 3-gram;
    * candidates outside the bucket match at most 2 bytes and can
      neither win nor influence the strict-improvement tie-break;
    * within a 4096-byte chunk ``copy_token_help``'s offset range
      always covers the whole chunk-so-far, so bucket entries never
      expire and ascending bucket order equals ascending scan order.

    The byte-exactness matters: Access validates the OVBA cache blob
    against its own compressor's output (see ``_encode_token_chunk``),
    and ``test_compress_byte_exact_against_access_sample_040`` plus the
    naive-oracle equivalence tests pin this property.
    """
    out = bytearray()
    pos = 0
    chunk_len = len(chunk)

    # 3-gram index: int key (little-endian packed three bytes) -> list
    # of positions in ascending order.  ``indexed_to`` marks how far the
    # index has been populated; positions are added lazily as the
    # encoder passes them, including positions inside copied regions.
    index: dict[int, list[int]] = {}
    indexed_to = 0

    while pos < chunk_len:
        flag_bits = 0
        tokens: list[bytes] = []

        for bit in range(8):
            if pos >= chunk_len:
                break

            length_mask, offset_mask, bit_count = copy_token_help(pos, 0)
            max_length = length_mask + 3

            best_len = 0
            best_offset = 0
            if pos + 3 <= chunk_len:
                while indexed_to < pos:
                    if indexed_to + 3 <= chunk_len:
                        key = (
                            chunk[indexed_to]
                            | chunk[indexed_to + 1] << 8
                            | chunk[indexed_to + 2] << 16
                        )
                        index.setdefault(key, []).append(indexed_to)
                    indexed_to += 1
                limit = max_length
                if chunk_len - pos < limit:
                    limit = chunk_len - pos
                target_key = chunk[pos] | chunk[pos + 1] << 8 | chunk[pos + 2] << 16
                for candidate in index.get(target_key, ()):
                    match_len = 3
                    while (match_len < limit
                           and chunk[candidate + match_len] == chunk[pos + match_len]):
                        match_len += 1
                    if match_len > best_len:
                        best_len = match_len
                        best_offset = pos - candidate
                        if best_len >= limit:
                            # No later candidate can strictly improve on
                            # the cap; the naive scan would keep this one.
                            break

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
    doc_string: str = ""             # MODULEDOCSTRING (MBCS)
    doc_string_unicode: str = ""     # MODULEDOCSTRINGUNICODE
    help_context: int = 0            # MODULEHELPCONTEXT
    cookie: int = 0                  # MODULECOOKIE (u16)


# ---------------------------------------------------------------------------
# Reference records (Gate 10)
# ---------------------------------------------------------------------------

@dataclass
class VBAReference:
    """A REFERENCED record from the dir stream's PROJECTREFERENCES section.

    [MS-OVBA] 2.3.4.2.2.  A reference is identified by one of several body
    record IDs; the optional ``libid`` strings carry registry/path data.
    """
    name: str = ""
    name_unicode: str = ""
    kind: str = ""              # "registered" | "project" | "control" | "original"
    libid: str = ""             # primary libid (registry path / file path)
    libid_secondary: str = ""   # twiddled / relative libid where applicable


def _prefer_unicode(ansi: str, unicode_value: str) -> str:
    """Return the authoritative value of an ANSI / UTF-16 dir record pair.

    The UTF-16 partner record is lossless by construction, while the ANSI
    record is limited to the project's code page.  When the two disagree
    the ANSI side is the damaged one -- either its bytes were undecodable
    (U+FFFD) or the writer had already substituted ``?`` for characters
    outside the page -- so the Unicode record wins (GitHub issue #12).
    Well-formed files carry identical values and are unaffected.
    """
    cleaned = unicode_value.rstrip("\x00")
    if cleaned and cleaned != ansi:
        return cleaned
    return ansi


def _parse_dir_stream(raw: bytes) -> tuple[_DirInfo, list[_ModuleInfo]]:
    """
    Parse a decompressed dir stream.

    Returns (_DirInfo, list[_ModuleInfo]).

    Decoding strategy: sequentially read records as (Id:u16, Size:u32, Data).
    PROJECTVERSION (0x0009) is the only record that lacks a real Size field
    (its Size slot is a Reserved=0x0004 marker and the payload is fixed at
    10 bytes including Reserved), so it is special-cased.  MODULENAME
    (0x0019) acts as a delimiter that starts a new _ModuleInfo entry once
    the parser has crossed into PROJECTMODULES.
    """
    pos = 0

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

    def _read(n: int) -> bytes:
        nonlocal pos
        chunk = raw[pos: pos + n]
        pos += n
        return chunk

    info = _DirInfo()
    modules: list[_ModuleInfo] = []
    current: _ModuleInfo | None = None
    in_modules = False
    pending_ref: VBAReference | None = None
    # After a REFERENCECONTROL TWIDDLED (0x002F) record, the spec
    # ([MS-OVBA] 2.3.4.2.2.9 / 2.3.4.2.2.7) allows a NameRecordExtended
    # (REFERENCENAME 0x0016 + optional 0x003E) inside the extended
    # record portion.  Those embedded name records must NOT be treated
    # as the start of a new reference.
    expecting_extended_name = False

    def _commit_ref() -> None:
        nonlocal pending_ref
        if pending_ref is not None:
            info.references.append(pending_ref)
            pending_ref = None

    def _enc() -> str:
        try:
            "".encode(f"cp{info.code_page}")
            return f"cp{info.code_page}"
        except LookupError:
            return "latin-1"

    while pos + 2 <= len(raw):
        record_id = _read_u16()

        # Module terminator (only meaningful after PROJECTMODULES).
        if record_id == 0x0010:   # dir terminator
            if pos + 4 <= len(raw):
                pos += 4  # consume trailing reserved u32
            break

        # PROJECTVERSION — special, no Size field.
        if record_id == 0x0009:
            if pos + 10 > len(raw):
                break
            _ = _read_u32()                 # Reserved=0x0004
            info.version_major = _read_u32()
            info.version_minor = _read_u16()
            continue

        # MODULENAME — starts a new module record.
        if record_id == 0x0019:
            in_modules = True
            _commit_ref()
            if current is not None:
                modules.append(current)
            current = _ModuleInfo(code_page=info.code_page)
            size = _read_u32()
            current.name = _read(size).decode(_enc(), errors="replace")
            continue

        if pos + 4 > len(raw):
            break
        size = _read_u32()
        data = _read(size)

        # ------------- PROJECTINFORMATION records -----------------
        if record_id == 0x0001:   # PROJECTSYSKIND
            if len(data) >= 4:
                info.sys_kind = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x004A: # PROJECTCOMPATVERSION
            if len(data) >= 4:
                info.compat_version = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x0002: # PROJECTLCID
            if len(data) >= 4:
                info.lcid = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x0014: # PROJECTLCIDINVOKE
            if len(data) >= 4:
                info.lcid_invoke = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x0003: # PROJECTCODEPAGE
            if len(data) >= 2:
                info.code_page = int(struct.unpack_from("<H", data, 0)[0])
        elif record_id == 0x0004: # PROJECTNAME
            info.name = data.decode(_enc(), errors="replace")
        elif record_id == 0x0005: # PROJECTDOCSTRING (MBCS)
            info.doc_string = data.decode(_enc(), errors="replace")
        elif record_id == 0x0040: # reserved partner: PROJECTDOCSTRINGUNICODE
            info.doc_string_unicode = data.decode("utf-16-le", errors="replace")
        elif record_id == 0x0006: # PROJECTHELPFILEPATH (MBCS)
            info.help_file = data.decode(_enc(), errors="replace")
        elif record_id == 0x003D: # reserved partner: PROJECTHELPFILEPATH2
            pass                  # duplicate of HelpFile, ignore
        elif record_id == 0x0007: # PROJECTHELPCONTEXT
            if len(data) >= 4:
                info.help_context = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x0008: # PROJECTLIBFLAGS
            if len(data) >= 4:
                info.lib_flags = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x000C: # PROJECTCONSTANTS (MBCS)
            info.constants = data.decode(_enc(), errors="replace")
        elif record_id == 0x003C: # reserved partner: PROJECTCONSTANTSUNICODE
            info.constants_unicode = data.decode("utf-16-le", errors="replace")
        elif record_id == 0x000F: # PROJECTMODULES (count)
            if len(data) >= 2:
                info.module_count = int(struct.unpack_from("<H", data, 0)[0])
            # First record of the PROJECTMODULES section — capture its absolute
            # start so save() can splice in a freshly serialized modules block.
            if info.modules_section_offset == -1:
                info.modules_section_offset = pos - 6 - len(data)
        elif record_id == 0x0013: # PROJECTCOOKIE
            if len(data) >= 2:
                info.project_cookie = int(struct.unpack_from("<H", data, 0)[0])

        # ------------- PROJECTREFERENCES records ------------------
        elif record_id == 0x0016: # REFERENCENAME (MBCS)
            if expecting_extended_name and pending_ref is not None:
                # Embedded name inside REFERENCECONTROL extended record.
                # Keep current pending_ref intact; do not start a new ref.
                pass
            else:
                _commit_ref()
                pending_ref = VBAReference(name=data.decode(_enc(), errors="replace"))
        elif record_id == 0x003E: # reserved partner: REFERENCENAME unicode
            if pending_ref is not None and not expecting_extended_name:
                pending_ref.name_unicode = data.decode("utf-16-le", errors="replace")
        elif record_id == 0x000D: # REFERENCEREGISTERED
            if pending_ref is None:
                pending_ref = VBAReference()
            pending_ref.kind = "registered"
            if len(data) >= 4:
                lib_size = int(struct.unpack_from("<I", data, 0)[0])
                pending_ref.libid = bytes(
                    data[4: 4 + lib_size]
                ).decode("latin-1", errors="replace")
            _commit_ref()
        elif record_id == 0x000E: # REFERENCEPROJECT
            if pending_ref is None:
                pending_ref = VBAReference()
            pending_ref.kind = "project"
            if len(data) >= 4:
                abs_size = int(struct.unpack_from("<I", data, 0)[0])
                pending_ref.libid = bytes(
                    data[4: 4 + abs_size]
                ).decode("latin-1", errors="replace")
                rel_off = 4 + abs_size
                if len(data) >= rel_off + 4:
                    rel_size = int(struct.unpack_from("<I", data, rel_off)[0])
                    pending_ref.libid_secondary = bytes(
                        data[rel_off + 4: rel_off + 4 + rel_size]
                    ).decode("latin-1", errors="replace")
            _commit_ref()
        elif record_id == 0x002F: # REFERENCECONTROL (twiddled libid)
            if pending_ref is None:
                pending_ref = VBAReference()
            pending_ref.kind = "control"
            if len(data) >= 4:
                tw_size = int(struct.unpack_from("<I", data, 0)[0])
                pending_ref.libid_secondary = bytes(
                    data[4: 4 + tw_size]
                ).decode("latin-1", errors="replace")
            # The TWIDDLED record is followed by an Extended record
            # ([MS-OVBA] 2.3.4.2.2.9) whose preamble is a NameRecordExtended
            # (REFERENCENAME + optional REFERENCENAME unicode partner).
            # Mark those as belonging to the current pending_ref.
            expecting_extended_name = True
            # do NOT commit yet: 0x0030 follow-up may contribute primary libid
        elif record_id == 0x0030: # REFERENCECONTROL extended (primary libid)
            if pending_ref is not None and len(data) >= 4:
                ex_size = int(struct.unpack_from("<I", data, 0)[0])
                pending_ref.libid = bytes(
                    data[4: 4 + ex_size]
                ).decode("latin-1", errors="replace")
            expecting_extended_name = False
            _commit_ref()
        elif record_id == 0x0033: # REFERENCEORIGINAL
            if pending_ref is None:
                pending_ref = VBAReference()
            pending_ref.kind = "original"
            pending_ref.libid = data.decode("latin-1", errors="replace")
            # Do not commit: REFERENCEORIGINAL is always followed by REFERENCECONTROL.

        # ------------- MODULE records -----------------------------
        # Each MBCS record is followed by its UTF-16 partner.  When the
        # MBCS decode was lossy -- the project's code page could not
        # represent the name, or the file declares a page we could not
        # resolve -- the Unicode record is the authoritative lossless
        # value, so prefer it (GitHub issue #12).
        elif record_id == 0x0047 and current is not None: # MODULENAMEUNICODE
            current.name_unicode = data.decode("utf-16-le", errors="replace")
            current.name = _prefer_unicode(current.name, current.name_unicode)
        elif record_id == 0x001A and current is not None: # MODULESTREAMNAME
            current.stream_name = data.decode(_enc(), errors="replace")
        elif record_id == 0x0032 and current is not None: # MODULESTREAMNAME unicode
            current.stream_name_unicode = data.decode("utf-16-le", errors="replace")
            current.stream_name = _prefer_unicode(
                current.stream_name, current.stream_name_unicode
            )
        elif record_id == 0x001C and current is not None: # MODULEDOCSTRING (MBCS)
            current.doc_string = data.decode(_enc(), errors="replace")
        elif record_id == 0x0048 and current is not None: # MODULEDOCSTRINGUNICODE
            current.doc_string_unicode = data.decode("utf-16-le", errors="replace")
            current.doc_string = _prefer_unicode(
                current.doc_string, current.doc_string_unicode
            )
        elif record_id == 0x0031 and current is not None: # MODULEOFFSET
            if len(data) >= 4:
                current.text_offset = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x001E and current is not None: # MODULEHELPCONTEXT
            if len(data) >= 4:
                current.help_context = int(struct.unpack_from("<I", data, 0)[0])
        elif record_id == 0x002C and current is not None: # MODULECOOKIE
            if len(data) >= 2:
                current.cookie = int(struct.unpack_from("<H", data, 0)[0])
        elif record_id == 0x0021 and current is not None: # MODULETYPE standard
            current.module_kind = VBAModuleKind.standard
        elif record_id == 0x0022 and current is not None: # MODULETYPE other
            current.module_kind = VBAModuleKind.other
        elif record_id == 0x0025 and current is not None: # MODULEREADONLY
            current.is_read_only = True
        elif record_id == 0x0028 and current is not None: # MODULEPRIVATE
            current.is_private = True
        elif record_id == 0x002B and current is not None: # Module terminator
            modules.append(current)
            current = None
        else:
            pass   # unknown record — already consumed via _read(size)

        _ = in_modules

    _commit_ref()
    if current is not None and current.name:
        modules.append(current)

    return info, modules


@dataclass
class _DirInfo:
    """Project-level fields recovered from the dir stream."""
    code_page: int = 1252
    name: str = ""
    sys_kind: int = 0
    compat_version: int = 0
    lcid: int = 0
    lcid_invoke: int = 0
    help_context: int = 0
    lib_flags: int = 0
    doc_string: str = ""
    doc_string_unicode: str = ""
    help_file: str = ""
    constants: str = ""
    constants_unicode: str = ""
    version_major: int = 0
    version_minor: int = 0
    module_count: int = 0
    project_cookie: int = 0
    # Absolute byte offset of the first PROJECTMODULES (0x000F) record within
    # the decompressed dir stream; -1 if the dir stream had no modules section.
    # Used by the writer to splice in a freshly serialized modules block while
    # leaving the project-information + references prefix verbatim.
    modules_section_offset: int = -1
    references: list[VBAReference] = field(default_factory=lambda: [])


# ---------------------------------------------------------------------------
# Public data model
# ---------------------------------------------------------------------------

class VBAModule:
    """A parsed VBA module with its source code.

    ``source`` may be materialized lazily: modules produced by
    :func:`parse_vba_project` defer the MS-OVBA decompression of their
    stream until the first ``.source`` access (issue #5 — decompression
    is 88-96% of the cost of opening a project, and callers that only
    want names, kinds, or one module out of many should not pay for all
    of it).  Constructing a module with an explicit ``source`` string is
    fully eager and behaves exactly as before.  One consequence of
    laziness: a corrupt chunk past the first one raises
    :class:`VBAProjectError` at first access rather than at parse time.
    """

    def __init__(
        self,
        name: str,
        stream_name: str,
        source: str = "",
        kind: VBAModuleKind = VBAModuleKind.standard,
        text_offset: int = 0,
        is_read_only: bool = False,
        is_private: bool = False,
        # Dir-stream sub-records preserved for round-trip serialization.
        name_unicode: str = "",
        stream_name_unicode: str = "",
        doc_string: str = "",
        doc_string_unicode: str = "",
        help_context: int = 0,
        cookie: int = 0,
        # Original bytes 0..text_offset of the module stream (performance
        # cache / version-dependent prefix).  Preserved across write-back
        # so that Office's cache invalidation logic operates the same way
        # as it would for an untouched stream.
        prefix_bytes: bytes = b"",
        # Whether the source has been edited and needs recompression on save.
        dirty: bool = False,
        # Cached attribute header captured at parse time so a body-only
        # source replacement (the VBE-style edit surface) can re-prepend
        # the required ``Attribute VB_*`` block.  Re-derived on demand when
        # missing via ``split_attribute_header(self.source)``.
        attribute_header: str = "",
        source_loader: Callable[[], str] | None = None,
    ) -> None:
        self.name = name
        self.stream_name = stream_name
        self.kind = kind
        self.text_offset = text_offset
        self.is_read_only = is_read_only
        self.is_private = is_private
        self.name_unicode = name_unicode
        self.stream_name_unicode = stream_name_unicode
        self.doc_string = doc_string
        self.doc_string_unicode = doc_string_unicode
        self.help_context = help_context
        self.cookie = cookie
        self.prefix_bytes = prefix_bytes
        self.dirty = dirty
        self.attribute_header = attribute_header
        if source_loader is not None:
            self._source: str | None = None
            self._source_loader: Callable[[], str] | None = source_loader
        else:
            self._source = source
            self._source_loader = None

    def __repr__(self) -> str:
        return (
            f"VBAModule(name={self.name!r}, stream_name={self.stream_name!r}, "
            f"kind={self.kind!r}, text_offset={self.text_offset}, "
            f"source_loaded={self._source is not None})"
        )

    @property
    def source_loaded(self) -> bool:
        """True once the source text is materialized in memory."""
        return self._source is not None

    @property
    def source(self) -> str:
        """The module's full source text (header plus body).

        Materializes lazily on first access for parsed modules.
        """
        if self._source is None:
            loader = self._source_loader
            assert loader is not None, "lazy module lost its loader"
            self._source = loader()
            self._source_loader = None
        return self._source

    @source.setter
    def source(self, value: str) -> None:
        self._source = value
        self._source_loader = None

    @property
    def body(self) -> str:
        """Return the source text with its leading attribute header stripped."""
        _, body = split_attribute_header(self.source)
        return body

    @body.setter
    def body(self, value: str) -> None:
        """Replace the body, preserving the cached ``attribute_header``."""
        header = self.attribute_header or split_attribute_header(self.source)[0]
        self.source = header + value
        self.attribute_header = header
        self.dirty = True

    def disassemble(self, *, is_64bit: bool = True) -> DisassembledModule:
        """Decode the compiled VBA7 p-code from this module's
        ``prefix_bytes`` (the binary PerformanceCache region containing
        the ``0xCAFE`` magic word).

        Returns an empty :class:`DisassembledModule` when the module
        has no p-code cached (e.g. modules that have never been
        compiled, or hosts that strip the cache on save).

        See :mod:`pyopenvba.vba_pcode` for the full opcode field guide.
        """
        from pyopenvba.vba_pcode import (
            DisassembledModule,
            disassemble_module_stream,
        )

        if not self.prefix_bytes:
            return DisassembledModule(
                cafe_offset=-1, num_lines=0, lines=()
            )
        return disassemble_module_stream(
            self.prefix_bytes, is_64bit=is_64bit
        )


@dataclass
class VBAProject:
    """Represents a parsed VBA project."""
    modules: list[VBAModule] = field(default_factory=lambda: [])
    code_page: int = 1252
    name: str = ""
    sys_kind: int = 0
    references: list[VBAReference] = field(default_factory=lambda: [])
    protection: ProjectProtection | None = None
    # Dir-stream PROJECTMODULES coupling fields.
    project_cookie: int = 0
    # Raw decompressed dir bytes captured at parse time, kept so that
    # save() can re-emit the project-information + references prefix verbatim
    # while regenerating only the modules section.  ``dir_modules_offset``
    # marks the start of the first 0x000F record within ``dir_raw``.
    # These fields are technically internal but are exposed without a leading
    # underscore so that pyopenvba.excel can drive the save path without
    # triggering pyright's reportPrivateUsage.
    dir_raw: bytes = field(default=b"", repr=False)
    dir_modules_offset: int = field(default=-1, repr=False)
    # Pending logical renames (old_name -> new_name); flushed by save().
    pending_renames: dict[str, str] = field(default_factory=lambda: {}, repr=False)
    # Pending newly-added stream names; flushed by save() into the CFB.
    pending_adds: set[str] = field(default_factory=lambda: set(), repr=False)
    # Pending stream names to delete from the CFB; flushed by save().
    pending_deletes: set[str] = field(default_factory=lambda: set(), repr=False)
    # Set whenever the module list's identity changes (add/rename/delete).
    dir_structure_dirty: bool = field(default=False, repr=False)

    def get_module(self, name: str) -> VBAModule:
        needle = name.casefold()
        for m in self.modules:
            if m.name.casefold() == needle:
                return m
        raise KeyError(f"Module not found: {name!r}")

    def module_names(self) -> list[str]:
        return [m.name for m in self.modules]

    # ------------------------------------------------------------------
    # Mutation API (Gate 13 / Gate 24)
    #
    # NOTE: these methods only update the in-memory project model.  The
    # write-back path (write_back_modules / save) emits only existing
    # module streams' source text — full dir/PROJECT stream regeneration
    # for added/renamed/deleted modules is not yet implemented; see
    # docs/roadmap.md.
    # ------------------------------------------------------------------

    def add_module(
        self,
        name: str,
        source: str,
        *,
        kind: VBAModuleKind = VBAModuleKind.standard,
        stream_name: str | None = None,
    ) -> VBAModule:
        """Append a new module to the project (in-memory only).

        ``source`` may be either a full source (already starting with an
        ``Attribute VB_*`` / ``VERSION ... CLASS`` header) or a bare body.
        For standard modules a minimal ``Attribute VB_Name = "<name>"``
        header is synthesized if one is not provided, matching the VBE UX
        where the user only types the executable body.

        For ``kind=other`` the source is normalized from file-export form
        to stream form first (``VERSION ... CLASS`` preamble stripped,
        ``Attribute VB_Base`` ensured), so ``.cls`` files exported from
        the VBE are accepted as-is.  See :func:`normalize_class_source`.

        Raises ``ValueError`` if a module with that name already exists.
        """
        needle = name.casefold()
        if any(m.name.casefold() == needle for m in self.modules):
            raise ValueError(f"Module already exists: {name!r}")

        # Normalize header: parse out any caller-supplied header, otherwise
        # synthesize a minimal standard-module header.  Class (kind=other)
        # sources are first converted from file-export form to stream form;
        # a bare body passes through unchanged and falls into the
        # synthesize_class_header branch below.
        if kind == VBAModuleKind.other:
            source = normalize_class_source(source)
        existing_header, _body = split_attribute_header(source)
        if existing_header:
            header = existing_header
            full_source = source
        elif kind == VBAModuleKind.standard:
            header = synthesize_standard_header(name)
            full_source = header + source
        elif kind == VBAModuleKind.other:
            header = synthesize_class_header(name)
            full_source = header + source
        else:
            raise ValueError(
                f"add_module(kind={kind.name}) requires the source to start "
                "with its own attribute header.  pyOpenVBA does not invent "
                "document or designer module headers."
            )

        module = VBAModule(
            name=name,
            stream_name=stream_name if stream_name is not None else name,
            source=full_source,
            kind=kind,
            text_offset=0,
            name_unicode=name,
            stream_name_unicode=(stream_name if stream_name is not None else name),
            attribute_header=header,
            dirty=True,
        )
        self.modules.append(module)
        self.dir_structure_dirty = True
        # If the same stream name was queued for deletion (delete-then-readd
        # of the same name, e.g. an idempotent "replace" script), cancel the
        # deletion and do NOT queue an add: the CFB stream and its PROJECT/
        # dir declaration are still on disk and must be reused as-is.  The
        # ``dirty=True`` flag above ensures the stream's source bytes are
        # rewritten via ``write_back_modules`` during ``save()``.  Without
        # this branch, ``save()`` would append a duplicate ``Module=NAME``
        # declaration (and matching ``[Workspace]`` entry) to PROJECT, which
        # Excel rejects as workbook corruption.
        if module.stream_name in self.pending_deletes:
            self.pending_deletes.discard(module.stream_name)
        else:
            self.pending_adds.add(module.stream_name)
        return module

    def rename_module(self, old_name: str, new_name: str) -> VBAModule:
        """Rename a module by logical name.

        Updates the in-memory project model and queues a CFB stream rename plus
        a dir/PROJECT stream rewrite that ``ExcelFile.save()`` will apply.
        """
        module = self.get_module(old_name)
        if module.name == new_name:
            return module
        needle = new_name.casefold()
        if needle != module.name.casefold() and any(
            m.name.casefold() == needle for m in self.modules
        ):
            raise ValueError(f"Module already exists: {new_name!r}")
        old_stream = module.stream_name
        old_logical = module.name
        module.name = new_name
        module.name_unicode = new_name
        # Keep stream name aligned with module name (matches Office behavior).
        module.stream_name = new_name
        module.stream_name_unicode = new_name
        # Re-key the in-source ``Attribute VB_Name = "..."`` line to the new
        # name.  Excel uses MODULENAME from the dir stream as the authoritative
        # binding, but VBE re-exports the attribute on edit and a stale name
        # in the source confuses round-trip diff tooling.  Only rewrite when
        # the existing header literally references the old logical name.
        if module.attribute_header:
            old_attr = f'Attribute VB_Name = "{old_logical}"'
            new_attr = f'Attribute VB_Name = "{new_name}"'
            if old_attr in module.attribute_header:
                module.attribute_header = module.attribute_header.replace(
                    old_attr, new_attr, 1
                )
                module.source = module.source.replace(old_attr, new_attr, 1)
        module.dirty = True
        self.dir_structure_dirty = True
        # Queue CFB stream rename only if the stream name actually changed.
        if old_stream != module.stream_name:
            # If this stream was queued as a pending add (never saved yet),
            # just retarget the add — no rename needs to hit the CFB.
            if old_stream in self.pending_adds:
                self.pending_adds.discard(old_stream)
                self.pending_adds.add(module.stream_name)
            else:
                # Compose with any prior pending rename so the chain collapses.
                existing_src = next(
                    (k for k, v in self.pending_renames.items() if v == old_stream),
                    None,
                )
                if existing_src is not None:
                    self.pending_renames[existing_src] = module.stream_name
                else:
                    self.pending_renames[old_stream] = module.stream_name
        # Also remember the logical-name change for PROJECT stream rewriting.
        # (We use the same dict keyed by logical names when stream==logical,
        # which is the common Excel case.)
        _ = old_logical
        return module

    def delete_module(self, name: str) -> None:
        """Remove a module from the project (in-memory only)."""
        module = self.get_module(name)
        self.modules.remove(module)
        self.dir_structure_dirty = True
        # If this was a never-saved module (still queued as an add), just
        # cancel the add; nothing exists on disk yet.
        if module.stream_name in self.pending_adds:
            self.pending_adds.discard(module.stream_name)
        else:
            self.pending_deletes.add(module.stream_name)
        # If a rename had been queued targeting this stream, drop it.
        for old, new in list(self.pending_renames.items()):
            if new == module.stream_name or old == module.stream_name:
                self.pending_renames.pop(old, None)

    # ------------------------------------------------------------------
    # Validation (Gate 19)
    # ------------------------------------------------------------------

    def validate(self, cfb: CFB | None = None) -> list[str]:
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

# Windows code-page identifiers whose portable Python codec is not
# spelled ``cp<N>``.
#
# These aliases are not merely a convenience: on Windows, CPython falls
# through to the operating system's code-page registry, so ``cp28592``
# and friends resolve there but raise LookupError on Linux and macOS.
# Relying on the ``cp<N>`` spelling therefore makes a project's text
# decode correctly on one platform and turn into latin-1 mojibake on
# another -- a platform-dependent data bug (caught by the cross-OS
# language-matrix CI job on its first run; see GitHub issue #13).
# Consulting this table first keeps behavior identical everywhere.
#
# Pages VBA hosts write that Python does spell ``cp<N>`` portably (874,
# 932, 936, 949, 950, 1250-1258, 1361, 65001) are deliberately absent.
_CODEPAGE_ALIASES: dict[int, str] = {
    # Macintosh
    10000: "mac_roman",
    10004: "mac_arabic",
    10006: "mac_greek",
    10007: "mac_cyrillic",
    10010: "mac_romanian",
    10029: "mac_latin2",
    10079: "mac_iceland",
    10081: "mac_turkish",
    10082: "mac_croatian",
    # KOI8
    20866: "koi8_r",
    21866: "koi8_u",
    # ISO 8859
    28591: "iso8859_1",
    28592: "iso8859_2",
    28593: "iso8859_3",
    28594: "iso8859_4",
    28595: "iso8859_5",
    28596: "iso8859_6",
    28597: "iso8859_7",
    28598: "iso8859_8",
    28599: "iso8859_9",
    28603: "iso8859_13",
    28605: "iso8859_15",
    # ISO 2022 / EUC / GB / misc
    50220: "iso2022_jp",
    50225: "iso2022_kr",
    51932: "euc_jp",
    51936: "gb2312",
    51949: "euc_kr",
    52936: "hz",
    54936: "gb18030",
    65000: "utf_7",
}


def _encoding_for_codepage(code_page: int) -> str:
    """Map a PROJECTCODEPAGE value to a Python codec name.

    Falls back to ``latin-1`` for pages Python cannot resolve, warning
    so the degradation is visible: silently decoding a foreign code page
    as latin-1 produces mojibake that survives round-trip checks (see
    GitHub issue #12).
    """
    alias = _CODEPAGE_ALIASES.get(code_page)
    if alias is not None:
        return alias
    encoding = f"cp{code_page}"
    try:
        codecs.lookup(encoding)
    except LookupError:
        warnings.warn(
            f"No Python codec for VBA code page {code_page}; falling back "
            "to latin-1.  Text outside the Latin-1 range will be mojibake. "
            "Please report this code page at "
            "https://github.com/WilliamSmithEdward/pyOpenVBA/issues",
            UserWarning,
            stacklevel=3,
        )
        return "latin-1"
    return encoding


def encode_mbcs(text: str, encoding: str) -> bytes:
    """Encode ``text`` to an MBCS code page, salvaging composed characters.

    Equivalent to ``text.encode(encoding, errors="replace")`` except for
    characters the codec cannot represent directly but *can* represent as
    a base character plus combining marks.  Python's charmap codecs do no
    composition, so Vietnamese text destined for cp1258 loses every
    stacked-diacritic character:

        'Tiếng Việt'.encode('cp1258', errors='replace')
        -> b'Ti?ng Vi?t'

    cp1258 stores ``ệ`` as precomposed ``ê`` (0xEA) plus a combining
    dot-below byte (0xF2), which is *not* the character's canonical
    decomposition, so NFD alone does not help either.  For each
    unmappable character this helper decomposes to NFD and tries folding
    each combining mark back into the base in turn, keeping the first
    combination the codec accepts and emitting the remaining marks as
    combining bytes.  Decoding those bytes yields a canonically
    equivalent (NFD-ish) string, so compare with NFC normalization.

    Characters that remain unmappable become ``?``, matching
    ``errors="replace"``.  Text that the codec already encodes directly
    is returned byte-for-byte unchanged, so this never alters output for
    projects that work today.
    """
    try:
        return text.encode(encoding)
    except UnicodeEncodeError:
        pass
    out = bytearray()
    for ch in text:
        try:
            out += ch.encode(encoding)
            continue
        except UnicodeEncodeError:
            pass
        decomposed = unicodedata.normalize("NFD", ch)
        if len(decomposed) > 1:
            base, marks = decomposed[0], list(decomposed[1:])
            folded: bytes | None = None
            for i, mark in enumerate(marks):
                combined = unicodedata.normalize("NFC", base + mark)
                if len(combined) != 1:
                    continue   # this mark does not compose with the base
                rest = "".join(marks[:i] + marks[i + 1:])
                try:
                    folded = combined.encode(encoding) + rest.encode(encoding)
                except UnicodeEncodeError:
                    continue
                break
            if folded is not None:
                out += folded
                continue
            try:
                out += decomposed.encode(encoding)
                continue
            except UnicodeEncodeError:
                pass
        out += ch.encode(encoding, errors="replace")
    return bytes(out)


def split_attribute_header(source: str) -> tuple[str, str]:
    """
    Split a VBA module source into ``(attribute_header, body)``.

    The header captures the optional class ``VERSION ... CLASS / BEGIN ... END``
    block, followed by any contiguous run of leading ``Attribute VB_*`` lines,
    and (if present) the single ``\\r\\n`` blank separator that VBE inserts
    between the attributes and the executable body.  Everything after that
    is the body.

    Detection is deterministic — header lines must literally start with
    ``Attribute `` or be part of the ``VERSION ... CLASS`` preamble.  No
    pattern-matching against names or values is performed.

    Always succeeds: when no header is present, returns ``("", source)``.
    """
    if not source:
        return "", ""
    # Normalize line ending lookup without mutating data.
    pos = 0
    n = len(source)

    # Optional VERSION ... CLASS block (class modules).
    if source.startswith("VERSION ") and " CLASS" in source[:64]:
        # Find the terminating "END\r\n" line.
        end_idx = source.find("\r\nEND\r\n")
        if end_idx != -1:
            pos = end_idx + len("\r\nEND\r\n")
        else:
            end_idx = source.find("\nEND\n")
            if end_idx != -1:
                pos = end_idx + len("\nEND\n")

    # Run of "Attribute " lines.
    while pos < n and source.startswith("Attribute ", pos):
        nl = source.find("\n", pos)
        if nl == -1:
            pos = n
            break
        pos = nl + 1

    # If no attribute lines and no VERSION block was consumed, no header.
    if pos == 0:
        return "", source

    # Consume a single blank-line separator if present (CRLF or LF).
    if source.startswith("\r\n", pos):
        pos += 2
    elif source.startswith("\n", pos):
        pos += 1

    return source[:pos], source[pos:]


def synthesize_standard_header(name: str) -> str:
    """Return the minimum attribute header for a new standard module."""
    return f'Attribute VB_Name = "{name}"\r\n'


# Universal CLSID for plain VBA class modules (not document/designer modules).
CLASS_MODULE_CLSID = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"


def synthesize_class_header(name: str) -> str:
    """Return the standard attribute header for a new plain class module.

    Uses the universal VBA class CLSID as ``VB_Base`` and the conventional
    default values for all other class attributes.  This header is equivalent
    to what VBE emits when you insert a new Class Module.

    Only suitable for plain class modules.  Document modules (ThisWorkbook,
    Sheet1, ThisDocument, etc.) and designer modules (UserForms) carry host-
    specific CLSIDs and must be provided with an explicit attribute header.
    """
    return (
        f'Attribute VB_Name = "{name}"\r\n'
        f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"\r\n'
        'Attribute VB_GlobalNameSpace = False\r\n'
        'Attribute VB_Creatable = False\r\n'
        'Attribute VB_PredeclaredId = False\r\n'
        'Attribute VB_Exposed = False\r\n'
        'Attribute VB_TemplateDerived = False\r\n'
        'Attribute VB_Customizable = False\r\n'
    )


def normalize_class_source(source: str, *, prior_header: str = "") -> str:
    """Convert a class module source from file-export form to stream form.

    The VBE writes two different textual forms of a class module.  The
    on-disk export (``.cls`` file) begins with a ``VERSION 1.0 CLASS`` /
    ``BEGIN`` ... ``END`` preamble and carries no ``Attribute VB_Base``
    line.  The form stored inside the module stream is the opposite: no
    VERSION preamble, and the attribute block includes ``VB_Base``.
    Writing export-form text into the stream produces a workbook that
    opens cleanly but fails to compile: a missing ``VB_Base`` raises
    "Invalid procedure call or argument" at the first ``New`` site, and
    a VERSION preamble is parsed as class body text and raises
    "Compile error: Expected: end of statement" (both verified against
    live Excel; see GitHub issue #1).

    Normalization steps:

    1. Strip a leading ``VERSION ... CLASS`` preamble if present.
    2. Ensure the attribute header carries a ``VB_Base`` line: a
       supplied one is kept untouched; otherwise the ``VB_Base`` line
       from ``prior_header`` (the module's existing stream header --
       this preserves host CLSIDs when a document module's source is
       replaced) is copied in; otherwise the universal class CLSID is
       inserted directly after ``Attribute VB_Name``.

    Idempotent: stream-form input is returned unchanged.  The input's
    line endings are preserved -- an inserted line copies the terminator
    of its anchor line.  Sources whose header has no ``VB_Name`` anchor
    (already invalid as a stream) are returned with only step 1 applied.
    """
    if not source:
        return source

    # Step 1: strip the export-only VERSION preamble.  Mirrors the
    # detection in split_attribute_header so the two functions always
    # agree on where a header begins.
    stripped = source
    if stripped.startswith("VERSION ") and " CLASS" in stripped[:64]:
        end_idx = stripped.find("\r\nEND\r\n")
        if end_idx != -1:
            stripped = stripped[end_idx + len("\r\nEND\r\n"):]
        else:
            end_idx = stripped.find("\nEND\n")
            if end_idx != -1:
                stripped = stripped[end_idx + len("\nEND\n"):]

    header, body = split_attribute_header(stripped)
    if not header:
        return stripped

    header_lines = header.splitlines(keepends=True)
    for line in header_lines:
        if line.startswith("Attribute VB_Base"):
            return stripped

    # Step 2: no VB_Base in the supplied header.  Find the VB_Name
    # anchor to insert after.
    insert_at = -1
    for i, line in enumerate(header_lines):
        if line.startswith("Attribute VB_Name"):
            insert_at = i + 1
            break
    if insert_at == -1:
        return stripped

    anchor = header_lines[insert_at - 1]
    if anchor.endswith("\r\n"):
        eol = "\r\n"
    elif anchor.endswith("\n"):
        eol = "\n"
    else:
        # Header ends without a newline; terminate the anchor so the
        # inserted line starts on its own line.
        eol = "\r\n"
        header_lines[insert_at - 1] = anchor + eol

    vb_base_line = ""
    if prior_header:
        for line in prior_header.splitlines():
            if line.startswith("Attribute VB_Base"):
                vb_base_line = line.rstrip("\r\n") + eol
                break
    if not vb_base_line:
        vb_base_line = f'Attribute VB_Base = "{CLASS_MODULE_CLSID}"{eol}'

    header_lines.insert(insert_at, vb_base_line)
    return "".join(header_lines) + body


def rebuild_module_stream(module: VBAModule, code_page: int) -> bytes:
    """
    Rebuild a module stream by preserving the original ``[0:text_offset]``
    performance-cache prefix and replacing ``[text_offset:]`` with a freshly
    compressed copy of ``module.source``.

    The replacement is byte-exact in the prefix region, so any cache-
    invalidation logic that Office performs on the prefix remains valid.
    """
    encoding = _encoding_for_codepage(code_page)
    source_bytes = encode_mbcs(module.source, encoding)
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
# dir-stream writer (Gate 13 persistence)
# ---------------------------------------------------------------------------

def _pack_record(rid: int, data: bytes) -> bytes:
    """[MS-OVBA] generic record encoding: Id(u16) + Size(u32) + Data."""
    return struct.pack("<HI", rid, len(data)) + data


def serialize_dir_modules_section(project: VBAProject) -> bytes:
    """
    Serialize the PROJECTMODULES section of a dir stream from ``project``.

    Output layout, per [MS-OVBA] 2.3.4.2.3:
        0x000F  PROJECTMODULES   Count:u16
        0x0013  PROJECTCOOKIE    Cookie:u16
        Module*  per-module record block
        0x0010  Terminator       (followed by Reserved:u32=0)

    Each per-module block emits:
        MODULENAME, MODULENAMEUNICODE, MODULESTREAMNAME, MODULESTREAMNAMEUNICODE,
        MODULEDOCSTRING, MODULEDOCSTRINGUNICODE, MODULEOFFSET, MODULEHELPCONTEXT,
        MODULECOOKIE, MODULETYPE, MODULEREADONLY?, MODULEPRIVATE?, Terminator.
    """
    enc = _encoding_for_codepage(project.code_page)
    out = bytearray()
    out += _pack_record(0x000F, struct.pack("<H", len(project.modules)))
    out += _pack_record(0x0013, struct.pack("<H", project.project_cookie & 0xFFFF))

    for m in project.modules:
        name = m.name
        name_u = m.name_unicode or m.name
        stream = m.stream_name or m.name
        stream_u = m.stream_name_unicode or stream
        out += _pack_record(0x0019, encode_mbcs(name, enc))
        out += _pack_record(0x0047, name_u.encode("utf-16-le"))
        out += _pack_record(0x001A, encode_mbcs(stream, enc))
        out += _pack_record(0x0032, stream_u.encode("utf-16-le"))
        out += _pack_record(0x001C, encode_mbcs(m.doc_string, enc))
        out += _pack_record(0x0048, m.doc_string_unicode.encode("utf-16-le"))
        out += _pack_record(0x0031, struct.pack("<I", m.text_offset))
        out += _pack_record(0x001E, struct.pack("<I", m.help_context))
        out += _pack_record(0x002C, struct.pack("<H", m.cookie & 0xFFFF))
        # MODULETYPE: 0x0021 standard / 0x0022 other (data is empty per spec).
        out += _pack_record(
            0x0021 if m.kind == VBAModuleKind.standard else 0x0022, b""
        )
        if m.is_read_only:
            out += _pack_record(0x0025, b"")
        if m.is_private:
            out += _pack_record(0x0028, b"")
        # Module terminator.
        out += _pack_record(0x002B, b"")

    # dir-stream terminator: 0x0010 with size 0, followed by Reserved:u32 = 0.
    out += _pack_record(0x0010, b"") + b"\x00\x00\x00\x00"
    return bytes(out)


# Public alias for cross-module use (e.g. pyopenvba.access).
# Kept alongside `serialize_dir_stream` so the read-side and write-side
# of the MS-OVBA dir stream are symmetric in the public surface.
parse_dir_stream = _parse_dir_stream


def serialize_dir_stream(project: VBAProject) -> bytes:
    """
    Build a fresh decompressed dir stream from ``project``.

    Splice strategy: the project-information + references prefix is reused
    verbatim from ``project._dir_raw`` (captured at parse time), and only the
    PROJECTMODULES section is regenerated.  This avoids round-trip drift in
    records the parser does not fully decode (e.g. extended REFERENCECONTROL
    cookies).
    """
    if project.dir_raw and project.dir_modules_offset >= 0:
        prefix = project.dir_raw[: project.dir_modules_offset]
    else:
        # Synthesized projects have no captured prefix; the caller is
        # responsible for supplying a valid PROJECTINFORMATION + REFERENCES
        # prefix in some other way before save.
        raise VBAProjectError(
            "serialize_dir_stream() requires dir_raw + dir_modules_offset; "
            "use a parsed VBAProject."
        )
    return prefix + serialize_dir_modules_section(project)


# ---------------------------------------------------------------------------
# PROJECT-stream writer (Gate 6 persistence)
# ---------------------------------------------------------------------------

def serialize_project_stream(
    raw: bytes,
    rename_map: dict[str, str],
    *,
    add_modules: list[tuple[str, str]] | None = None,
    delete_names: set[str] | None = None,
    code_page: int = 1252,
) -> bytes:
    """
    Rewrite a PROJECT stream's plain-text body to apply pending mutations.

    ``rename_map``:
        Old-logical-name to new-logical-name.  Rewrites the right-hand side of
        ``Module=`` / ``Class=`` / ``BaseClass=`` / ``Document=NAME/&H...``
        and keys inside ``[Workspace]``.
    ``add_modules``:
        Sequence of ``(name, declaration_key)`` pairs where ``declaration_key``
        is either ``"Module"`` (standard) or ``"Class"`` (class / other).
        Each pair is appended as ``"<key>=<name>"`` after the existing
        declarations and a matching ``[Workspace]`` entry is appended.
    ``delete_names``:
        Logical names to remove.  Any ``Module=NAME`` / ``Class=NAME`` /
        ``BaseClass=NAME`` / ``Document=NAME/...`` line with a matching name
        is dropped, as is any ``[Workspace]`` key with that name.

    Everything else (``ID``, ``Name``, ``CMG``, ``DPB``, ``GC``,
    ``[Host Extender Info]``) is preserved byte-for-byte except for the
    targeted substitutions.

    ``code_page``:
        PROJECTCODEPAGE of the owning project.  The PROJECT stream is
        code-page ANSI per [MS-OVBA] 2.3.1, so its ``Module=`` /
        ``Document=`` declarations carry module names in that page.
        Defaults to 1252 for callers that have no project handy.

    Returns code-page-encoded bytes with CRLF line endings.
    """
    add_modules = add_modules or []
    delete_names = delete_names or set()
    # NOTE: we deliberately do not short-circuit on "all inputs empty" —
    # callers may invoke this to run the dedup/normalization pass on a
    # PROJECT stream that was previously corrupted by buggy writes (e.g.
    # duplicate ``Module=NAME`` lines from a delete-then-readd flow).
    encoding = _encoding_for_codepage(code_page)
    text = raw.decode(encoding, errors="replace")

    rename_ci: dict[str, str] = {k.casefold(): v for k, v in rename_map.items()}
    delete_ci: set[str] = {n.casefold() for n in delete_names}

    # Track which declaration names have already been emitted (case-folded)
    # so that any duplicate ``Module=``/``Class=``/``BaseClass=``/``Document=``
    # lines that may already exist in the input PROJECT stream are scrubbed
    # on the way out.  Excel treats duplicate declarations as workbook
    # corruption, so save() must converge to a deduplicated state even when
    # repairing an already-broken input.
    seen_decls: set[str] = set()
    seen_workspace: set[str] = set()

    out_lines: list[str] = []
    in_workspace = False
    # Track the last index where a module-declaration line was emitted, so
    # we can splice newly-added Module=/Class= lines right after them.
    last_decl_idx = -1
    # Track where [Workspace] block sits so we can append new entries to it.
    workspace_idx = -1
    workspace_end_idx = -1

    for line in text.splitlines():
        stripped = line.strip()
        if stripped.lower() == "[workspace]":
            in_workspace = True
            workspace_idx = len(out_lines)
            out_lines.append(line)
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            if in_workspace:
                workspace_end_idx = len(out_lines)
            in_workspace = False
            out_lines.append(line)
            continue
        if "=" not in stripped:
            out_lines.append(line)
            continue
        key, _, value = stripped.partition("=")
        key_s = key.strip()
        if in_workspace:
            if key_s.casefold() in delete_ci:
                continue
            new_key = rename_ci.get(key_s.casefold(), key_s)
            ws_key = new_key.casefold()
            if ws_key in seen_workspace:
                continue
            seen_workspace.add(ws_key)
            out_lines.append(f"{new_key}={value.strip()}")
            continue
        if key_s in ("Module", "Class", "BaseClass"):
            v = value.strip()
            if v.casefold() in delete_ci:
                continue
            new_val = rename_ci.get(v.casefold(), v)
            decl_key = new_val.casefold()
            if decl_key in seen_decls:
                continue
            seen_decls.add(decl_key)
            out_lines.append(f"{key_s}={new_val}")
            last_decl_idx = len(out_lines) - 1
            continue
        if key_s == "Document":
            name_part, sep, id_part = value.strip().partition("/")
            if name_part.casefold() in delete_ci:
                continue
            new_name = rename_ci.get(name_part.casefold(), name_part)
            decl_key = new_name.casefold()
            if decl_key in seen_decls:
                continue
            seen_decls.add(decl_key)
            out_lines.append(f"Document={new_name}{sep}{id_part}")
            last_decl_idx = len(out_lines) - 1
            continue
        out_lines.append(line)

    # Splice in any newly-added module declarations after the last existing
    # one; if there were none, append at the project-section end (immediately
    # before the first '[' header, or at the very end).  Skip any name that
    # already has a declaration (defensive: caller may pass an add for a
    # module whose declaration survived from a previous save).
    fresh_adds = [
        (name, decl_key)
        for name, decl_key in add_modules
        if name.casefold() not in seen_decls
    ]
    if fresh_adds:
        insert_at = last_decl_idx + 1 if last_decl_idx >= 0 else _project_section_end(out_lines)
        new_decl_lines = [f"{decl_key}={name}" for name, decl_key in fresh_adds]
        out_lines[insert_at:insert_at] = new_decl_lines
        for name, _ in fresh_adds:
            seen_decls.add(name.casefold())
        # Shift workspace markers if they were after the insertion point.
        if workspace_idx >= insert_at:
            workspace_idx += len(new_decl_lines)
        if workspace_end_idx >= insert_at:
            workspace_end_idx += len(new_decl_lines)

        # Append matching [Workspace] entries.  Real Office writes
        # "<Name>=0, 0, 0, 0, C" by default; replicate that.  Skip names
        # that already have a Workspace entry.
        if workspace_idx >= 0:
            ws_new = [
                (name, f"{name}=0, 0, 0, 0, C")
                for name, _ in fresh_adds
                if name.casefold() not in seen_workspace
            ]
            if ws_new:
                ws_insert = workspace_end_idx if workspace_end_idx > workspace_idx else len(out_lines)
                out_lines[ws_insert:ws_insert] = [line for _, line in ws_new]
                for name, _ in ws_new:
                    seen_workspace.add(name.casefold())

    return encode_mbcs("\r\n".join(out_lines) + "\r\n", encoding)


def _project_section_end(lines: list[str]) -> int:
    """Return the index of the first ``[...]`` header line, or ``len(lines)``."""
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith("[") and s.endswith("]"):
            return i
    return len(lines)


# ---------------------------------------------------------------------------
# Public factory
# ---------------------------------------------------------------------------

def _make_source_loader(
    compressed: bytes, encoding: str, stream_name: str
) -> Callable[[], str]:
    """Build the deferred decompress-and-decode thunk for a lazy module."""
    def _load() -> str:
        return decompress(compressed, stream_name=stream_name).decode(
            encoding, errors="replace"
        )
    return _load


def parse_vba_project(cfb: CFB) -> VBAProject:
    """
    Extract all VBA module metadata from a parsed CFB.

    The CFB must be the vbaProject.bin from an xlsm/xlsb, or the whole
    file for an xls workbook.

    Module source text is loaded lazily: only the first compressed
    chunk of each module stream is decompressed here (enough for the
    ``Attribute VB_*`` header, and for single-chunk modules it is
    already the whole source).  The remaining chunks decompress on the
    first ``VBAModule.source`` access.  Stream lookup and MODULEOFFSET
    bounds checks stay eager so structural corruption still surfaces at
    parse time.
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

    dir_raw = decompress(dir_compressed, stream_name="VBA/dir")
    dir_info, module_infos = _parse_dir_stream(dir_raw)
    code_page = dir_info.code_page
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

        compressed_source = bytes(stream_compressed[info.text_offset:])
        ovba_name = f"VBA/{stream_name}"

        # Decompress just the first chunk.  For single-chunk modules
        # (decompressed size <= 4096) this is the complete source and
        # the module is materialized eagerly with zero extra work.
        prefix_raw, consumed_all = _decompress(
            compressed_source, stream_name=ovba_name, max_bytes=1
        )
        prefix_text = prefix_raw.decode(encoding, errors="replace")

        source: str | None
        loader: Callable[[], str] | None
        if consumed_all:
            source = prefix_text
            header, _ = split_attribute_header(prefix_text)
            loader = None
        else:
            header, body_prefix = split_attribute_header(prefix_text)
            if header and body_prefix:
                # Header fits inside the first chunk (the normal case);
                # defer the rest of the stream.
                source = None
                loader = _make_source_loader(
                    compressed_source, encoding, ovba_name
                )
            else:
                # Either no attribute header at all, or the header runs
                # to the chunk boundary and may be truncated.  Both are
                # abnormal; fall back to eager full decompression so
                # attribute_header is always derived from complete text.
                source = decompress(
                    compressed_source, stream_name=ovba_name
                ).decode(encoding, errors="replace")
                header, _ = split_attribute_header(source)
                loader = None

        modules.append(VBAModule(
            name=info.name,
            stream_name=stream_name,
            source=source if source is not None else "",
            kind=info.module_kind,
            text_offset=info.text_offset,
            is_read_only=info.is_read_only,
            is_private=info.is_private,
            name_unicode=info.name_unicode,
            stream_name_unicode=info.stream_name_unicode,
            doc_string=info.doc_string,
            doc_string_unicode=info.doc_string_unicode,
            help_context=info.help_context,
            cookie=info.cookie,
            prefix_bytes=bytes(stream_compressed[: info.text_offset]),
            attribute_header=header,
            source_loader=loader,
        ))

    project = VBAProject(modules=modules, code_page=code_page)
    project.name = dir_info.name
    project.sys_kind = dir_info.sys_kind
    project.references = dir_info.references
    project.project_cookie = dir_info.project_cookie
    project.dir_raw = dir_raw
    project.dir_modules_offset = dir_info.modules_section_offset

    # Attach protection state from the PROJECT stream if present.
    try:
        project_stream_raw = cfb.get_stream("PROJECT")
    except KeyError:
        project_stream_raw = None
    if project_stream_raw is not None:
        try:
            ps = parse_project_stream(project_stream_raw, code_page=code_page)
            project.protection = ps.protection
        except VBAProjectError:
            pass

    return project


# ---------------------------------------------------------------------------
# PROJECT stream parser ([MS-OVBA] 2.3.1)
# ---------------------------------------------------------------------------

@dataclass
class ProjectProtection:
    """State of the [MS-OVBA] 2.3.1 PROJECT-stream protection records.

    The actual ``CMG``/``DPB``/``GC`` values are XOR-obfuscated, so we
    only expose presence and the raw obfuscated hex strings here.  Code
    that needs to read the underlying password hash must do the XOR
    de-obfuscation itself.
    """
    has_password: bool = False
    cmg: str = ""          # raw obfuscated DocLockProj/HostProj/VbeProj bits
    dpb: str = ""          # raw obfuscated password hash
    gc: str = ""           # raw obfuscated DocLockProj-visible bit


@dataclass
class ProjectStream:
    """Decoded view of the plain-text PROJECT stream."""
    items: list[tuple[str, str]] = field(default_factory=lambda: [])
    workspace: dict[str, str] = field(default_factory=lambda: {})
    standard_modules: list[str] = field(default_factory=lambda: [])
    class_modules: list[str] = field(default_factory=lambda: [])
    document_modules: list[tuple[str, str]] = field(default_factory=lambda: [])  # (name, &HID)
    base_classes: list[str] = field(default_factory=lambda: [])
    name: str = ""
    id: str = ""
    protection: ProjectProtection = field(default_factory=lambda: ProjectProtection())
    host_extender_info: list[str] = field(default_factory=lambda: [])


def parse_project_stream(raw: bytes, *, code_page: int = 1252) -> ProjectStream:
    """Parse the plain-text PROJECT stream.

    The PROJECT stream is code-page ANSI key=value text terminated by
    CRLF ([MS-OVBA] 2.3.1); pass the project's PROJECTCODEPAGE as
    ``code_page`` so non-Latin module names decode correctly.  Empty
    lines separate the project information block, the ``[Host Extender
    Info]`` block, and the ``[Workspace]`` block.
    """
    text = raw.decode(_encoding_for_codepage(code_page), errors="replace")

    out = ProjectStream()
    section = "project"  # "project" | "host_extender" | "workspace"

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.lower() == "[host extender info]":
            section = "host_extender"
            continue
        if stripped.lower() == "[workspace]":
            section = "workspace"
            continue
        if section == "host_extender":
            out.host_extender_info.append(stripped)
            continue

        if "=" not in stripped:
            continue
        key, _, value = stripped.partition("=")
        key = key.strip()
        value = value.strip()
        # Strip surrounding quotes when present.
        if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
            value = value[1:-1]

        if section == "workspace":
            out.workspace[key] = value
            continue

        out.items.append((key, value))

        # Module classification keys.
        if key == "Module":
            out.standard_modules.append(value)
        elif key == "Class":
            out.class_modules.append(value)
        elif key == "BaseClass":
            out.base_classes.append(value)
        elif key == "Document":
            # Value form is "Name/&H00000000..."
            name_part, _, id_part = value.partition("/")
            out.document_modules.append((name_part, id_part))
        elif key == "Name":
            out.name = value
        elif key == "ID":
            out.id = value
        elif key == "CMG":
            out.protection.cmg = value
        elif key == "DPB":
            out.protection.dpb = value
            # Per MS-OVBA 2.3.1.16, DPB is present even on unprotected projects
            # but the de-obfuscated payload's first byte indicates protection
            # state.  Treat any DPB longer than a trivial placeholder as
            # password-bearing.  A trivial empty/placeholder hash decodes to
            # ~30 hex chars; real passwords push DPB above ~60 chars.
            out.protection.has_password = len(value) >= 60
        elif key == "GC":
            out.protection.gc = value

    return out


# ---------------------------------------------------------------------------
# PROJECTwm stream parser ([MS-OVBA] 2.3.4.4)
# ---------------------------------------------------------------------------

def parse_projectwm(
    raw: bytes, *, code_page: int = 1252
) -> list[tuple[str, str]]:
    """Parse the PROJECTwm stream into (mbcs_name, unicode_name) pairs.

    ``code_page`` selects the MBCS codec for the first name in each pair;
    pass the project's PROJECTCODEPAGE so non-Latin names decode
    correctly.

    Format:
        loop:
            ModuleName_MBCS    : null-terminated bytes
            if empty: break (single 0x00 terminator)
            ModuleName_Unicode : null-terminated UTF-16-LE (terminator u16=0)
    """
    out: list[tuple[str, str]] = []
    pos = 0
    n = len(raw)
    while pos < n:
        # Find next null byte to delimit the MBCS name.
        nul = raw.find(b"\x00", pos)
        if nul == -1:
            break
        mbcs = raw[pos:nul]
        pos = nul + 1
        if not mbcs:
            # Terminator entry.
            break
        # Read UTF-16-LE name terminated by a zero u16.
        start = pos
        while pos + 1 < n:
            if raw[pos] == 0 and raw[pos + 1] == 0:
                break
            pos += 2
        unicode_name = raw[start:pos].decode("utf-16-le", errors="replace")
        pos += 2   # skip the u16=0 terminator
        mbcs_name = mbcs.decode(
            _encoding_for_codepage(code_page), errors="replace"
        )
        out.append((mbcs_name, unicode_name))
    return out


def serialize_projectwm(
    pairs: Iterable[tuple[str, str]], *, code_page: int = 1252
) -> bytes:
    """Serialize ``(mbcs_name, unicode_name)`` pairs to a PROJECTwm stream.

    The stream format ([MS-OVBA] 2.3.4.4) is a sequence of records each
    consisting of a null-terminated MBCS module name followed by a
    null-terminated (u16=0) UTF-16-LE module name.  The stream is
    terminated by a single empty MBCS record (one 0x00 byte).

    ``code_page`` selects the MBCS code page used to encode the first
    name in each pair; unencodable characters are replaced.
    """
    enc = _encoding_for_codepage(code_page)
    buf = bytearray()
    for mbcs, unicode_name in pairs:
        if not mbcs:
            # An empty MBCS name would terminate the stream prematurely.
            continue
        buf += encode_mbcs(mbcs, enc)
        buf += b"\x00"
        buf += unicode_name.encode("utf-16-le", errors="replace")
        buf += b"\x00\x00"
    # Final terminator per [MS-OVBA] 2.3.4.4: empty MBCS name (single 0x00)
    # plus a trailing alignment byte (Excel always emits two zero bytes).
    buf += b"\x00\x00"
    return bytes(buf)


# ---------------------------------------------------------------------------
# PROJECTlk stream parser ([MS-OVBA] 2.3.4.5)
# ---------------------------------------------------------------------------

@dataclass
class LicenseRecord:
    """A LicenseInfoRecord from the PROJECTlk stream."""
    lic_key: bytes = b""
    libid: str = ""
    classid: bytes = b""        # 16-byte GUID
    cookie: int = 0


def parse_projectlk(raw: bytes) -> list[LicenseRecord]:
    """Parse the PROJECTlk stream into LicenseRecord entries.

    Each LicenseInfoRecord:
        Id:u16 = 0x0001
        Size:u32
        SizeOfLicKey:u32
        LicKey:bytes[SizeOfLicKey]
        LibidSize:u32
        Libid:bytes[LibidSize]
        ClassID:bytes[16]
        Cookie:u32
    """
    out: list[LicenseRecord] = []
    pos = 0
    n = len(raw)
    while pos + 6 <= n:
        rec_id = int(struct.unpack_from("<H", raw, pos)[0])
        size = int(struct.unpack_from("<I", raw, pos + 2)[0])
        pos += 6
        if rec_id != 0x0001 or pos + size > n:
            break
        body = raw[pos: pos + size]
        pos += size

        rec = LicenseRecord()
        bpos = 0
        if bpos + 4 <= len(body):
            lic_size = int(struct.unpack_from("<I", body, bpos)[0])
            bpos += 4
            if bpos + lic_size <= len(body):
                rec.lic_key = bytes(body[bpos: bpos + lic_size])
                bpos += lic_size
        if bpos + 4 <= len(body):
            lib_size = int(struct.unpack_from("<I", body, bpos)[0])
            bpos += 4
            if bpos + lib_size <= len(body):
                rec.libid = bytes(body[bpos: bpos + lib_size]).decode(
                    "latin-1", errors="replace"
                )
                bpos += lib_size
        if bpos + 16 <= len(body):
            rec.classid = bytes(body[bpos: bpos + 16])
            bpos += 16
        if bpos + 4 <= len(body):
            rec.cookie = int(struct.unpack_from("<I", body, bpos)[0])
        out.append(rec)
    return out


def serialize_projectlk(records: Iterable[LicenseRecord]) -> bytes:
    """Serialize ``LicenseRecord`` entries to a PROJECTlk stream.

    Inverse of :func:`parse_projectlk`.  Each record is emitted as a
    LicenseInfoRecord (Id=0x0001) per [MS-OVBA] 2.3.4.5.  ClassID is
    zero-padded or truncated to 16 bytes.
    """
    out = bytearray()
    for rec in records:
        classid = rec.classid[:16].ljust(16, b"\x00")
        libid_bytes = rec.libid.encode("latin-1", errors="replace")
        body = bytearray()
        body += struct.pack("<I", len(rec.lic_key))
        body += rec.lic_key
        body += struct.pack("<I", len(libid_bytes))
        body += libid_bytes
        body += classid
        body += struct.pack("<I", rec.cookie & 0xFFFFFFFF)
        out += struct.pack("<HI", 0x0001, len(body))
        out += body
    return bytes(out)


# ---------------------------------------------------------------------------
# V3 content hash (Gate 15)
# ---------------------------------------------------------------------------

def compute_v3_content_hash(project: VBAProject) -> bytes:
    """Compute a stable SHA-1 digest over the project's source content.

    NOTE: this is a pyOpenVBA-internal canonical digest, not the exact
    Office VBE V3 content hash used for digital signatures (which depends
    on host-specific normalization rules and tokenization that are
    outside the MS-OVBA spec).  It is stable across read-modify-write
    round trips that don't change source text.
    """
    import hashlib
    h = hashlib.sha1()
    h.update(struct.pack("<I", project.code_page))
    for m in sorted(project.modules, key=lambda x: x.name.casefold()):
        h.update(m.name.casefold().encode("utf-8"))
        h.update(b"\0")
        # Normalize line endings so CRLF/LF round trips are stable.
        normalized = m.source.replace("\r\n", "\n").replace("\r", "\n")
        h.update(normalized.encode("utf-8", errors="replace"))
        h.update(b"\0")
    return h.digest()


# ---------------------------------------------------------------------------
# Signature detection (Gate 17)
# ---------------------------------------------------------------------------

@dataclass
class SignatureInfo:
    """Summary of digital-signature streams found inside a VBA CFB."""
    present: bool = False
    kinds: list[str] = field(default_factory=lambda: [])   # "legacy" | "agile" | "v3"


_SIGNATURE_STREAM_NAMES = {
    "_VBA_PROJECT_SIGNATURE":       "legacy",
    "_VBA_PROJECT_SIGNATURE_AGILE": "agile",
    "_VBA_PROJECT_SIGNATURE_V3":    "v3",
}


def invalidate_vba_project_cache(cfb: CFB) -> bool:
    """Invalidate the ``_VBA_PROJECT`` performance cache.

    Per [MS-OVBA] 2.3.4.1 the ``_VBA_PROJECT`` stream is laid out as::

        Reserved1 (2 bytes) = 0x61CC
        Version   (2 bytes) = host-defined cookie
        Reserved2 (1 byte)  = 0x00
        PerformanceCache (variable bytes; "undefined and MUST be ignored on read")

    The PerformanceCache may reference offsets / module identities that
    become stale after a mutating save (add / rename / delete / source
    edit).  Office regenerates the cache on next open, but if the cache
    body still parses, some Office builds have been observed to honor
    its stale contents and either re-prompt for repair or surface
    incorrect p-code on first open.

    This helper preserves the 5-byte header (so the host still
    recognises the stream) and zeroes the cache body in place.  Returns
    ``True`` if the stream was found and invalidated, ``False`` if the
    stream is absent or too short to carry a header.
    """
    try:
        stream = cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")
    except KeyError:
        return False
    if len(stream) < 5:
        return False
    if stream[0] != 0xCC or stream[1] != 0x61:
        # Not the documented header; do not touch.
        return False
    if stream[4] != 0x00:
        return False
    if len(stream) == 5:
        return False  # Already minimal; nothing to invalidate.
    body = bytes(stream[:5]) + bytes(len(stream) - 5)
    cfb.write_stream_in_storage("VBA", "_VBA_PROJECT", body)
    return True


def detect_signature(cfb: CFB) -> SignatureInfo:
    """Detect VBA digital-signature streams in the given CFB.

    Returns a :class:`SignatureInfo` whose ``present`` is True if any of
    the three known signature streams are present inside the
    ``_VBA_PROJECT_CUR/VBA`` storage path (xlsm/xlsb) or the root
    (legacy xls embedding).
    """
    info = SignatureInfo()
    candidates: list[str] = []
    try:
        candidates.extend(cfb.list_streams_in_storage("VBA"))
    except KeyError:
        pass
    try:
        candidates.extend(cfb.list_streams())
    except KeyError:
        pass
    for name in candidates:
        kind = _SIGNATURE_STREAM_NAMES.get(name)
        if kind is not None and kind not in info.kinds:
            info.kinds.append(kind)
            info.present = True
    return info
