"""Access ``.accdb`` write primitives: resize an LVAL row and rebuild a
module's compiled p-code region.

This is the storage half of the Access write path. The code-generation
half lives in :mod:`vba_compile`.

Three things have to agree before Access will load a modified module, and
missing any one of them is what defeated the earlier phase-5 attempts:

1. **The LVAL row itself.** Rows grow downward from the end of a 4 KB
   page; the slot table at ``+14`` holds their offsets and the page's
   free-space counter lives at ``+2``. Growing a row means shifting every
   row physically below it and fixing all three.

2. **The MSysAccessStorage length field.** Every VBA long-value is
   described by a catalog row holding ``<u16 length> 00 40 <slot><page>``.
   Access trusts that length over the slot table, so a resized row whose
   catalog length still reads the old value makes Access fault while
   loading the project. This is the "internal index binding a module to
   its storage" that section 5 of the lessons document listed as never
   located.

3. **The dir stream's MODULEOFFSET.** It records where the compressed
   source begins inside the module row, so it moves whenever the p-code
   before it changes size.

The OVBA compressor here emits literal-only chunks. The repository's
general :func:`pyopenvba.vba.compress` produces smaller output that our
own decompressor accepts, but Access rejected it in testing; literal-only
chunks carry no copy tokens and load correctly.

Dev-only research code: Windows, desktop Access and ``pyvbaharness`` are
needed to verify anything it produces.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "docs"
                       / "research" / "pcode"))
from pcode_hash import identifier_hash

from pyopenvba.access_read import AccessReader
from pyopenvba.vba import decompress

ACE_PAGE_SIZE = 4096
STORAGE_PAGE_DEFAULT = 48

# Two u16 fields in the module-stream header count the lines of the
# procedure region. They track the line table's size, and Access refuses
# to load a module whose copies are stale -- which is what made an early
# attempt at changing the statement count fail. Both hold the same value.
PROC_LINE_COUNT_OFFSETS = (516, 518)

# MS-OVBA dir record MODULEOFFSET: <u16 id=0x0031><u32 size=4><u32 value>
_MODULEOFFSET_HDR = bytes.fromhex("310004000000")

# The project identifier table lives in the CC 61 _VBA_PROJECT row. It
# opens with the 'Access' reference record and closes with this sentinel.
# Immediately BEFORE the table sit two u16 counters that Access validates:
# the number of identifier records, and the highest slot in use. Appending
# a record without bumping both makes Access hang while loading.
_TABLE_HEADS = (b"\x02\x00\x06\x0cAccess", b"\x02\x00\x06\x04Access")
_TABLE_SENTINEL = b"\x02\xff\xff\x01\x01"
_COUNT_BACKOFF = 10          # u16 identifier count, at table_start - 10
_SLOT_COUNT_BACKOFF = 12     # u16 slot count,       at table_start - 12


# --- MS-OVBA compression, literal-only ---------------------------------
def compress_literal_only(data: bytes) -> bytes:
    """Compress ``data`` into MS-OVBA chunks that use no copy tokens.

    Each chunk is flagged compressed (so it may be shorter than the 4096
    bytes a raw chunk must carry) but every token is a literal, which is
    the encoding Access accepts most readily. A chunk holds at most 3640
    decompressed bytes so its 1-flag-byte-per-8-literals overhead still
    fits the 4096-byte chunk-data limit.
    """
    out = bytearray([0x01])
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 3640]
        pos += len(chunk)
        body = bytearray()
        for i in range(0, len(chunk), 8):
            body.append(0x00)          # 8 literal flags
            body += chunk[i:i + 8]
        header = 0x8000 | 0x3000 | ((len(body) - 1) & 0x0FFF)
        out += header.to_bytes(2, "little")
        out += body
    return bytes(out)


# --- LVAL page / row surgery -------------------------------------------
def slot_offsets(data: bytearray, base: int) -> list[int]:
    n = int.from_bytes(data[base + 12:base + 14], "little")
    return [int.from_bytes(data[base + 14 + 2 * i:base + 16 + 2 * i], "little")
            for i in range(n)]


def row_extent(data: bytearray, base: int, slot: int) -> tuple[int, int]:
    """Byte range of one row within its page: it ends at the next-higher
    live slot offset, or at the end of the page."""
    slots = slot_offsets(data, base)
    start, end = slots[slot] & 0x0FFF, ACE_PAGE_SIZE
    for other in slots:
        if (other & 0xF000) == 0xD000:      # tombstone
            continue
        off = other & 0x0FFF
        if start < off < end:
            end = off
    return start, end


def _set_slot(data: bytearray, base: int, slot: int, offset: int) -> None:
    cur = int.from_bytes(data[base + 14 + 2 * slot:base + 16 + 2 * slot],
                         "little")
    data[base + 14 + 2 * slot:base + 16 + 2 * slot] = (
        ((cur & 0xF000) | (offset & 0x0FFF)).to_bytes(2, "little"))


def write_row(data: bytearray, page: int, slot: int, payload: bytes) -> int:
    """Replace one LVAL row, resizing its page in place.

    Rows below the target shift by the size delta, their slot offsets are
    rewritten, and the page free-space counter at ``+2`` is corrected.
    Returns the delta. Raises if the page cannot absorb the growth.
    """
    base = page * ACE_PAGE_SIZE
    start, end = row_extent(data, base, slot)
    delta = len(payload) - (end - start)
    if delta == 0:
        data[base + start:base + end] = payload
        return 0
    slots = slot_offsets(data, base)
    live = [(i, s & 0x0FFF) for i, s in enumerate(slots)
            if (s & 0xF000) != 0xD000]
    lowest = min(off for _, off in live)
    table_end = 14 + 2 * len(slots)
    if lowest - delta < table_end:
        raise ValueError(
            f"page {page} cannot grow slot {slot} by {delta} bytes "
            f"(free={lowest - table_end})")
    block = bytes(data[base + lowest:base + start])
    data[base + lowest - delta:base + start - delta] = block
    data[base + start - delta:base + end] = payload
    for i, off in live:
        if off <= start:
            _set_slot(data, base, i, off - delta)
    free = int.from_bytes(data[base + 2:base + 4], "little")
    data[base + 2:base + 4] = (free - delta).to_bytes(2, "little")
    return delta


def set_storage_length(data: bytearray, lval_page: int, lval_slot: int,
                       new_length: int,
                       storage_page: int = STORAGE_PAGE_DEFAULT) -> int:
    """Update the MSysAccessStorage length for one VBA long-value.

    The catalog row ends with ``<u16 length> 00 40 <slot><page>``; we find
    it by its unique page/slot pointer and rewrite the length in front.
    Returns the previous value.
    """
    base = storage_page * ACE_PAGE_SIZE
    ptr = bytes([0x00, 0x40, lval_slot & 0xFF, lval_page & 0xFF])
    hits = []
    i = data.find(ptr, base, base + ACE_PAGE_SIZE)
    while i >= 0:
        hits.append(i)
        i = data.find(ptr, i + 1, base + ACE_PAGE_SIZE)
    if len(hits) != 1:
        raise ValueError(f"storage pointer {ptr.hex()} matched {len(hits)} "
                         "rows, expected exactly 1")
    pos = hits[0]
    old = int.from_bytes(data[pos - 2:pos], "little")
    data[pos - 2:pos] = new_length.to_bytes(2, "little")
    return old


def find_project_row(path) -> tuple[int, int, bytes]:
    """Locate the ``CC 61`` _VBA_PROJECT row: ``(page, slot, bytes)``."""
    for page, slot, row in AccessReader(path)._iter_lval_rows():
        if bytes(row).startswith(b"\xcc\x61"):
            return page, slot, bytes(row)
    raise ValueError(f"no _VBA_PROJECT row found in {path}")


def append_identifiers(row: bytes, names, code_page: int = 1252) -> bytes:
    """Append identifier records to the project table in ``row``.

    Each record is ``<u8 len><u8 type=0><name><u16 hash><0x10 0x00>``,
    where the hash is the OLE ``LHashValOfNameSysA`` value that Access
    itself stores (see ``docs/research/pcode/pcode_hash.py``). The two
    u16 counters ahead of the table are bumped to match; Access hangs on
    load if they disagree with the records.

    Appending keeps every existing record's position, so p-code operands
    already in the module (``524 + 2*index``) stay valid. A new name
    takes the next index.
    """
    start = -1
    for head in _TABLE_HEADS:
        start = row.find(head)
        if start >= 0:
            break
    if start < 0:
        raise ValueError("identifier table not found in _VBA_PROJECT row")
    sentinel = row.find(_TABLE_SENTINEL, start)
    if sentinel < 0:
        raise ValueError("identifier table sentinel not found")

    encoding = f"cp{code_page}"
    records = bytearray()
    for name in names:
        body = name.encode(encoding)
        if not 0 < len(body) < 64:
            raise ValueError(f"identifier {name!r} has an unusable length")
        records += bytes((len(body), 0x00)) + body
        records += identifier_hash(name, code_page=encoding).to_bytes(2, "little")
        records += b"\x10\x00"

    out = bytearray(row[:sentinel] + bytes(records) + row[sentinel:])
    for backoff in (_COUNT_BACKOFF, _SLOT_COUNT_BACKOFF):
        off = start - backoff
        value = int.from_bytes(out[off:off + 2], "little")
        out[off:off + 2] = (value + len(names)).to_bytes(2, "little")
    return bytes(out)


def find_dir_row(path) -> tuple[int, int, bytes, bytes]:
    """Locate the MS-OVBA dir stream: ``(page, slot, decompressed, raw)``."""
    reader = AccessReader(path)
    for page, slot, row in reader._iter_lval_rows():
        try:
            raw = decompress(bytes(row))
        except Exception:
            continue
        if raw[:6] == b"\x01\x00\x04\x00\x00\x00":
            return page, slot, raw, bytes(row)
    raise ValueError(f"no dir stream found in {path}")


def module_name_positions(dir_stream: bytes) -> list[tuple[int, str]]:
    """``(record_position, name)`` for each MODULENAME record."""
    out: list[tuple[int, str]] = []
    for match in re.finditer(rb"\x19\x00(....)", dir_stream, re.S):
        size = int.from_bytes(match.group(1), "little")
        if not 0 < size < 64:
            continue
        name = dir_stream[match.end():match.end() + size]
        if len(name) == size and all(0x20 <= c < 0x7F for c in name):
            out.append((match.start(), name.decode("latin-1")))
    return out


def find_moduleoffset_pos(dir_stream: bytes, module: str | None = None) -> int:
    """Byte position of a module's MODULEOFFSET *value*.

    A project holds one MODULEOFFSET per module. Each belongs to the
    MODULENAME record that precedes it, which is how ``module`` selects
    one; with ``module=None`` a single-module project is assumed.
    """
    hits = [m.start() for m in re.finditer(re.escape(_MODULEOFFSET_HDR),
                                           dir_stream)]
    if not hits:
        raise ValueError("no MODULEOFFSET record in dir stream")
    if module is None:
        if len(hits) != 1:
            names = [n for _, n in module_name_positions(dir_stream)]
            raise ValueError(
                f"{len(hits)} modules in this project ({', '.join(names)}); "
                "pass module=<name> to choose one")
        return hits[0] + 6
    names = module_name_positions(dir_stream)
    for start in hits:
        owner = [n for pos, n in names if pos < start]
        if owner and owner[-1].lower() == module.lower():
            return start + 6
    raise ValueError(f"module {module!r} has no MODULEOFFSET record")


# --- the module stream's compiled p-code region ------------------------
def align8(value: int) -> int:
    return (value + 7) & ~7


class Perf:
    """The ``0xCAFE`` performance-cache region of a module stream.

    Layout, measured against Access-compiled modules (2026-08)::

        [0 : cafe]              header; u32 @29 = absolute end of p-code
        [cafe]                  FE CA
        [cafe+4]                u16 line count
        [cafe+6 : +12*n]        12-byte line records
        [.. : +10]              10-byte gap; u16 @+6 = total p-code size
        [pstart : +total]       p-code, then an 8-byte trailer
        [end_pcode : modoff]    6 fixed bytes
        [modoff :]              OVBA-compressed source

    A 12-byte line record is ``<flags><80|81><08|09><indent>`` then a u16
    p-code length at ``+4``, a u16 frame-size hint at ``+6``, and a u32
    p-code offset at ``+8``. Line offsets are 8-byte aligned; a line with
    offset ``0xFFFFFFFF`` carries no p-code. ``indent`` is the source
    line's leading-space count.
    """

    def __init__(self, row: bytes, moduleoffset: int) -> None:
        self.row = bytes(row)
        self.modoff = moduleoffset
        self.cafe = self.row.find(b"\xfe\xca")
        off = self.cafe + 4
        self.num_lines = int.from_bytes(self.row[off:off + 2], "little")
        off += 2
        self.rec_start = off
        self.recs = [bytearray(self.row[off + 12 * i:off + 12 * (i + 1)])
                     for i in range(self.num_lines)]
        off += 12 * self.num_lines
        self.gap = bytearray(self.row[off:off + 10])
        self.pstart = off + 10
        self.total = int.from_bytes(self.gap[6:8], "little")
        self.end_pcode = self.pstart + self.total
        self.middle = self.row[self.end_pcode:self.modoff]
        self.src_comp = self.row[self.modoff:]
        self.lines: list[bytes | None] = []
        for rec in self.recs:
            length = int.from_bytes(rec[4:6], "little")
            offset = int.from_bytes(rec[8:12], "little")
            self.lines.append(
                None if (offset == 0xFFFFFFFF or length <= 0)
                else bytes(self.row[self.pstart + offset:
                                    self.pstart + offset + length]))
        last_end = 0
        for rec, code in zip(self.recs, self.lines, strict=True):
            if code is None:
                continue
            last_end = max(last_end,
                           int.from_bytes(rec[8:12], "little") + len(code))
        self.trailer = bytes(self.row[self.pstart + align8(last_end):
                                      self.end_pcode])

    def source(self) -> bytes:
        return decompress(self.src_comp)

    def build(self, new_lines=None, new_source=None, lines=None,
              recs=None) -> tuple[bytes, int]:
        """Re-emit the module row, recomputing every dependent field.

        ``new_lines`` patches individual lines by index; ``lines``/``recs``
        replace the tables wholesale (which changes the line count).
        Returns ``(row_bytes, new_moduleoffset)``.
        """
        out_lines = list(self.lines) if lines is None else list(lines)
        for index, code in (new_lines or {}).items():
            out_lines[index] = code
        out_recs = ([bytearray(r) for r in self.recs] if recs is None
                    else [bytearray(r) for r in recs])

        buf = bytearray()
        for index, code in enumerate(out_lines):
            if code is None:
                continue
            offset = align8(len(buf))
            buf += bytes(offset - len(buf))
            out_recs[index][8:12] = offset.to_bytes(4, "little")
            out_recs[index][4:6] = len(code).to_bytes(2, "little")
            buf += code
        buf += bytes(align8(len(buf)) - len(buf))
        buf += self.trailer
        total = len(buf)

        gap = bytearray(self.gap)
        gap[6:8] = total.to_bytes(2, "little")
        source = (self.src_comp if new_source is None
                  else compress_literal_only(new_source))

        out = bytearray(self.row[:self.rec_start])
        out[self.cafe + 4:self.cafe + 6] = len(out_recs).to_bytes(2, "little")
        for rec in out_recs:
            out += rec
        out += gap
        out += buf
        # The record table may have resized, so the p-code start moves.
        end_pcode = self.rec_start + 12 * len(out_recs) + 10 + total
        assert len(out) == end_pcode, (len(out), end_pcode)
        out += self.middle
        new_modoff = len(out)
        out += source
        out[29:33] = end_pcode.to_bytes(4, "little")
        # Keep the procedure line counters in step with the line table.
        # They are relative-patched rather than recomputed: the absolute
        # value counts one procedure's lines, so only the delta is known
        # to be right for a module holding several.
        delta = len(out_recs) - self.num_lines
        if delta:
            for off in PROC_LINE_COUNT_OFFSETS:
                if off + 2 > self.cafe:
                    continue
                value = int.from_bytes(self.row[off:off + 2], "little")
                out[off:off + 2] = (value + delta).to_bytes(2, "little")
        return bytes(out), new_modoff


def load_module(path, module: str | None = None) -> dict:
    """Collect everything needed to rewrite one module in ``path``.

    ``module`` names the module to target; the default takes the only
    one, and raises if the project holds several.
    """
    reader = AccessReader(path)
    streams = reader.find_module_streams()
    if not streams:
        raise ValueError(f"{path}: no VBA module with compiled p-code")
    if module is None:
        stream = streams[0]
    else:
        matches = [s for s in streams if s.name.lower() == module.lower()]
        if not matches:
            raise ValueError(
                f"module {module!r} not found; have "
                f"{', '.join(s.name for s in streams)}")
        stream = matches[0]
    dir_page, dir_slot, dir_dec, _ = find_dir_row(path)
    pos = find_moduleoffset_pos(dir_dec, stream.name if module else module)
    return {"page": stream.page, "slot": stream.slot,
            "row": bytes(stream.raw), "dir_page": dir_page,
            "dir_slot": dir_slot, "dir_dec": dir_dec,
            "modoff": int.from_bytes(dir_dec[pos:pos + 4], "little"),
            "modoff_pos": pos}


def write_module(data: bytearray, info: dict, new_row: bytes,
                 new_modoff: int) -> None:
    """Splice a rebuilt module row into a database image.

    Updates, in the order Access needs them consistent: the module row and
    its catalog length, then the dir stream's MODULEOFFSET and *its*
    catalog length.
    """
    write_row(data, info["page"], info["slot"], new_row)
    set_storage_length(data, info["page"], info["slot"], len(new_row))
    dir_dec = bytearray(info["dir_dec"])
    pos = info["modoff_pos"]
    dir_dec[pos:pos + 4] = new_modoff.to_bytes(4, "little")
    new_dir = compress_literal_only(bytes(dir_dec))
    write_row(data, info["dir_page"], info["dir_slot"], new_dir)
    set_storage_length(data, info["dir_page"], info["dir_slot"], len(new_dir))
