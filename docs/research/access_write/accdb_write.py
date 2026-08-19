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

Compression uses the repository's :func:`pyopenvba.vba.compress`. An
earlier round of this work concluded Access rejected it and fell back to
literal-only chunks, but that test was confounded by a stale catalog
length; with the length correct, Access accepts it. It matters: on one
module the real compressor produced 148 bytes where literal-only produced
366, matching Access's own output exactly. ``compress_literal_only`` is
kept for comparison.

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

from pyopenvba.access_read import PAGE_TYPE_DATA, AccessReader
from pyopenvba.vba import compress, decompress

ACE_PAGE_SIZE = 4096
STORAGE_PAGE_DEFAULT = 48

# Two u16 fields in the module-stream header count the lines of the
# procedure region. They track the line table's size, and Access refuses
# to load a module whose copies are stale -- which is what made an early
# attempt at changing the statement count fail. Both hold the same value.
PROC_LINE_COUNT_OFFSETS = (516, 518)

# FuncDefn opens a procedure and EndFunc closes it; the opcode is the low
# 10 bits of a line's first word, and FuncDefn's u32 operand follows it.
FUNCDEFN_OPCODE = 150
ENDFUNC_OPCODE = 105

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
        # The u16 in the gap holds the p-code region's size, but it is
        # only 16 bits wide and a large module overflows it -- a 65544
        # byte region stores as 8. The u32 at offset 29 records the same
        # region's end and does not overflow, so trust that and keep the
        # u16 as the low half it is.
        self.end_pcode = int.from_bytes(self.row[29:33], "little")
        self.total = self.end_pcode - self.pstart
        if self.total < 0 or (self.total & 0xFFFF) != int.from_bytes(
                self.gap[6:8], "little"):
            # Fall back if the two disagree in a way overflow cannot explain.
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
        self.counter_base = find_counter_base(self.row, self.lines,
                                              self.num_lines, self.cafe)

    def source(self) -> bytes:
        return decompress(self.src_comp)

    def source_lines(self) -> list[str]:
        """Source split so index *i* lines up with line-table entry *i*.

        The leading ``Attribute`` block is not represented in the line
        table, and it is not one line: a standard module carries only
        ``VB_Name`` while a class module carries five. Counting the block
        instead of assuming its size is what keeps class modules aligned.
        """
        lines = self.source().decode("latin-1").split("\r\n")
        start = 0
        while start < len(lines) and lines[start].startswith("Attribute "):
            start += 1
        return lines[start:]

    def attribute_lines(self) -> list[str]:
        """The leading ``Attribute`` block, which has no line records."""
        lines = self.source().decode("latin-1").split("\r\n")
        start = 0
        while start < len(lines) and lines[start].startswith("Attribute "):
            start += 1
        return lines[:start]

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
        gap[6:8] = (total & 0xFFFF).to_bytes(2, "little")
        source = (self.src_comp if new_source is None
                  else compress(new_source))

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
        _write_procedure_line_counts(out, out_lines, len(out_recs),
                                     self.cafe, self.counter_base)
        return bytes(out), new_modoff


def _procedure_line_counts(lines, num_lines: int) -> list[tuple[int, int]]:
    """``(func_ operand, expected counter)`` for each procedure, in order.

    A procedure owns every source line from its ``FuncDefn`` up to the
    next procedure's ``FuncDefn``, or to the end of the module for the
    last one -- so the blank separator line between two procedures counts
    toward the one above it. Measured against Access's own output over a
    controlled series (one procedure, one to six body lines) and on two-
    and three-procedure modules; matches 97 of the 103 modules in this
    repo that have a procedure at all, the six exceptions being fixtures
    whose p-code was deliberately left inconsistent with their source.

    ``EndFunc`` deliberately plays no part: a ``Declare`` emits a
    ``FuncDefn`` with no matching ``EndFunc``, and pairing them up drops
    its counter and shifts every later one.
    """
    starts: list[tuple[int, int]] = []
    for index, code in enumerate(lines):
        if not code or len(code) < 6:
            continue
        if int.from_bytes(code[:2], "little") & 0x03FF == FUNCDEFN_OPCODE:
            starts.append((int.from_bytes(code[2:6], "little"), index))
    out: list[tuple[int, int]] = []
    for position, (func_operand, start) in enumerate(starts):
        following = starts[position + 1][1] if position + 1 < len(starts) else num_lines
        out.append((func_operand, max(0, following - start)))
    return out


def find_counter_base(row: bytes, lines, num_lines: int, cafe: int) -> int | None:
    """Offset the per-procedure line counters are measured from.

    Each procedure's pair sits at ``base + func_``. The base is 516 for
    an ordinary standard module, but not universally -- a class module
    was measured at 612, with its first procedure's ``func_`` starting at
    56 rather than 0. Rather than collect constants, locate the base by
    finding the one offset at which every procedure's stored pair already
    equals the value the layout implies.
    """
    wanted = _procedure_line_counts(lines, num_lines)
    if not wanted:
        return None
    candidates = []
    for base in range(0, min(cafe, 4096) - 4, 2):
        for func_operand, value in wanted:
            off = base + func_operand
            if off + 4 > cafe:
                break
            if int.from_bytes(row[off:off + 2], "little") != value:
                break
            if int.from_bytes(row[off + 2:off + 4], "little") != value:
                break
        else:
            candidates.append(base)
    return candidates[0] if len(candidates) == 1 else None


def _write_procedure_line_counts(out: bytearray, lines, num_lines: int,
                                 cafe: int, base: int | None) -> None:
    """Refresh the per-procedure line counters.

    Every procedure owns a pair of u16 counters at ``base + func_`` and
    ``base + func_ + 2``, where ``func_`` is its FuncDefn operand and the
    base is 516 for a standard module and 612 for a class module. Each
    holds the number of source lines the procedure spans -- see
    :func:`_procedure_line_counts`.

    These are computed outright rather than shifted by the module's line
    delta, which is what lets a procedure other than the first be
    rewritten: editing a later one then leaves the earlier counters alone,
    exactly as Access does.
    """
    if base is None:
        return
    for func_operand, value in _procedure_line_counts(lines, num_lines):
        off = base + func_operand
        if off + 4 <= cafe:
            out[off:off + 2] = value.to_bytes(2, "little")
            out[off + 2:off + 4] = value.to_bytes(2, "little")


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
    # A module too large for one page is stored as a chain of rows, with
    # `stream.raw` holding the assembled chain. Reading that is fine;
    # writing it back is not, so record it and let write_module refuse.
    head = bytes(reader._lval_row_bytes(stream.page, stream.slot))
    chained = len(stream.raw) != len(head)
    dir_page, dir_slot, dir_dec, dir_raw = find_dir_row(path)
    pos = find_moduleoffset_pos(dir_dec, stream.name if module else module)
    return {"page": stream.page, "slot": stream.slot,
            "row": bytes(stream.raw), "dir_page": dir_page,
            "dir_slot": dir_slot, "dir_dec": dir_dec,
            "modoff": int.from_bytes(dir_dec[pos:pos + 4], "little"),
            "modoff_pos": pos, "chained": chained, "name": stream.name,
            "dir_raw": dir_raw}


def write_module(data: bytearray, info: dict, new_row: bytes,
                 new_modoff: int) -> None:
    """Splice a rebuilt module row into a database image.

    Updates, in the order Access needs them consistent: the module row and
    its catalog length, then the dir stream's MODULEOFFSET and *its*
    catalog length.
    """
    set_lval_payload(data, info["page"], info["slot"], new_row,
                     len(info["row"]))
    dir_dec = bytearray(info["dir_dec"])
    pos = info["modoff_pos"]
    dir_dec[pos:pos + 4] = new_modoff.to_bytes(4, "little")
    new_dir = compress(bytes(dir_dec))
    set_lval_payload(data, info["dir_page"], info["dir_slot"], new_dir,
                     len(info["dir_raw"]))


# --- declarations: Dim and its header record ---------------------------
# `Dim x As Long` is eight bytes of p-code -- `Dim | VarDefn(var_)`, the
# same eight whatever the type -- plus a 24-byte record in the pre-0xCAFE
# header that carries everything else. Without the record Access crashes;
# with the wrong type it silently miscompiles, so both halves are needed.
#
# Records live at ``DECL_BASE + var_``, with ``var_`` starting at 88 and
# striding by 24, and form a linked list. Two shapes, by whether another
# declaration follows:
#
#   not last  TT ffff 0000 0000 8460 next ffff ffff frame ffff ffff ffff
#   last      TT ffff 0000 0000 ffff ffff 0000 0000 8302 owner ffff ffff
#
# `TT` is the VARTYPE code, `next` the following declaration's name
# operand, `owner` the module's, and `frame` the variable's frame offset
# (-40 for the first, eight less each time, so every local takes eight
# bytes whatever its type). The last record has no frame field: that slot
# carries the owner link instead.
DECL_BASE = 464
DECL_FIRST_VAR = 88
DECL_RECORD = 24
DECL_TABLE_GAP = 48
DECL_BUCKETS = 16
_DECL_NEXT_TAG = 0x8460
_DECL_OWNER_TAG = 0x8302
_U16_NULL = 0xFFFF
_U32_NULL = 0xFFFFFFFF

# VBA's VARTYPE codes, the same numbering the Coerce op_types use.
VARTYPE = {"integer": 2, "long": 3, "single": 4, "double": 5,
           "currency": 6, "date": 7, "string": 8, "object": 9,
           "boolean": 11, "variant": 12, "byte": 17}

# Fields to correct after inserting a record. Measured identical across
# S1->T2, T2->T3 and the D-series: "abs" offsets sit before the insertion
# point, "rel" offsets are measured from it.
_DECL_FIXUPS_ABS = ((9, 24), (25, 24), (444, 24), (488, 24), (540, 24),
                    (492, -8))
_DECL_FIXUPS_REL = ((32, 24), (156, 24), (220, 24), (228, 24), (256, 24),
                    (260, 24), (296, -8), (316, -24), (392, 24), (446, 24))

# Appending is exact while the module has at most three declarations, and
# stops being exact at four. Measured across three independent name
# series (aa/bb/cc..., pp/qq/zz..., and a set chosen to collide in one
# bucket): every 1->2, 2->3 and 3->4 transition reproduces Access byte
# for byte, and every 4->5 differs by the same eight bytes in a second
# structure -- one holding Variant-typed records -- that reorganizes at
# five entries and is not yet characterized. Refused rather than guessed,
# on the same principle as `_require_reproducible`.
MAX_MODELLED_DECLARATIONS = 3


def declaration_count(header: bytes) -> int:
    """How many declarations the module already has."""
    count = 0
    while True:
        off = DECL_BASE + DECL_FIRST_VAR + DECL_RECORD * count
        if off + DECL_RECORD > len(header):
            return count
        tag = int.from_bytes(header[off + 16:off + 18], "little")
        following = int.from_bytes(header[off + 8:off + 10], "little")
        if tag == _DECL_OWNER_TAG:
            return count + 1
        if following != _DECL_NEXT_TAG:
            return count
        count += 1


def _decl_words(fields) -> bytes:
    return b"".join(int(f & 0xFFFF).to_bytes(2, "little") for f in fields)


def _bump(out: bytearray, offset: int, delta: int, size: int = 4) -> None:
    """Add ``delta``, leaving an all-ones null sentinel alone.

    All-ones means "no value", not a number: incrementing it turns a null
    pointer into offset 23, and Access crashes on the result.
    """
    null = _U16_NULL if size == 2 else _U32_NULL
    value = int.from_bytes(out[offset:offset + size], "little")
    if value == null:
        return
    out[offset:offset + size] = ((value + delta) & null).to_bytes(size, "little")


def add_declaration(header: bytes, name: str, vartype: int, name_operand: int,
                    *, line_delta: int = 1) -> bytes:
    """Append one declaration to a module header.

    ``name_operand`` is the new variable's entry in the project identifier
    table, which must already exist. Returns the grown header; the caller
    shifts ``MODULEOFFSET`` by :data:`DECL_RECORD` to match.
    """
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pcode"))
    from pcode_hash import identifier_hash

    count = declaration_count(header)
    if count > MAX_MODELLED_DECLARATIONS:
        raise ValueError(
            f"module already has {count} declarations; appending past "
            f"{MAX_MODELLED_DECLARATIONS} does not reproduce Access "
            "byte-for-byte, so it is refused rather than guessed")
    insert = DECL_BASE + DECL_FIRST_VAR + DECL_RECORD * count
    out = bytearray(header[:insert]) + bytearray(DECL_RECORD)         + bytearray(header[insert:])
    previous = insert - DECL_RECORD
    owner = (int.from_bytes(header[previous + 18:previous + 20], "little")
             if count else 0)
    out[insert:insert + DECL_RECORD] = _decl_words(
        [vartype, _U16_NULL, 0, 0, _U16_NULL, _U16_NULL, 0, 0,
         _DECL_OWNER_TAG, owner, _U16_NULL, _U16_NULL])
    # The name's bucket in the local-name hash table, which is the same
    # OLE LHashValOfNameSysA the project identifier table uses. The table
    # holds procedures as well as variables, so a new name can displace
    # one either way.
    table = (DECL_BASE + DECL_FIRST_VAR + DECL_RECORD * (count + 1)
             + DECL_TABLE_GAP)
    # A bucket holding a *procedure* record moves with the insertion; one
    # holding a variable does not, because var_ offsets are fixed. Tell
    # them apart by whether the value is one of the existing var_ slots.
    variables = {DECL_FIRST_VAR + DECL_RECORD * k for k in range(count + 1)}
    for bucket in range(DECL_BUCKETS):
        cell = table + 4 * bucket
        value = int.from_bytes(out[cell:cell + 4], "little")
        if value != _U32_NULL and value not in variables:
            out[cell:cell + 4] = (value + DECL_RECORD).to_bytes(4, "little")
    # Records carrying a displaced *procedure* offset move with it too.
    for k in range(count):
        field = DECL_BASE + DECL_FIRST_VAR + DECL_RECORD * k + 12
        value = int.from_bytes(out[field:field + 4], "little")
        if value not in (_U32_NULL, 0) and value not in variables:
            out[field:field + 4] = (value + DECL_RECORD).to_bytes(4, "little")
    slot = table + 4 * (identifier_hash(name) % DECL_BUCKETS)
    # Read the bucket in the *shifted* header, before overwriting it.
    displaced = int.from_bytes(out[slot:slot + 4], "little")
    out[slot:slot + 4] = (DECL_FIRST_VAR + DECL_RECORD * count).to_bytes(
        4, "little")
    if count:
        # The record before it stops being last, gains a frame offset, and
        # takes custody of whatever this name displaced from its bucket --
        # its own offset when a variable lost the bucket, the procedure's
        # when a procedure did, and the null marker when it was empty.
        kept = int.from_bytes(header[previous:previous + 2], "little")
        out[previous:previous + DECL_RECORD] = _decl_words(
            [kept, _U16_NULL, 0, 0, _DECL_NEXT_TAG, name_operand,
             displaced & 0xFFFF, (displaced >> 16) & 0xFFFF,
             (0xFFD8 - 8 * (count - 1)) & 0xFFFF,
             _U16_NULL, _U16_NULL, _U16_NULL])
    for offset, delta in _DECL_FIXUPS_ABS:
        _bump(out, offset, delta)
    for offset, delta in _DECL_FIXUPS_REL:
        _bump(out, insert + offset, delta)
    for offset in (516, 518):
        _bump(out, offset, line_delta, size=2)
    return bytes(out)


# --- the __SRP_* compiled-code cache -----------------------------------
# Access keeps a second, compiled form of every module in storage rows
# named ``__SRP_0``, ``__SRP_1``, ... -- the same performance cache that
# [MS-OVBA] describes as ``__SRP_*`` streams, and that ``_host.py``
# already drops when writing Office files. Access *executes* that cache,
# so a module rewritten in the canonical p-code keeps its old behaviour
# until the cache is gone. Dropping these rows is what makes a written
# module take effect; Access then recompiles from what we wrote.
#
# The rows are catalogued in ``MSysAccessStorage``. Deleting only the
# long-value rows leaves the catalog pointing at nothing and Access
# rejects the whole project ("can't find the function"), so the catalog
# row is what must go.
_SRP_NAME = "__SRP_".encode("utf-16-le")
_ROW_DELETED = 0x8000


def drop_srp_cache(data: bytearray) -> int:
    """Delete every ``__SRP_*`` catalog row; return how many were dropped.

    Marks the row's slot-table entry deleted, which is how Access itself
    retires a row. The long-value rows the entries pointed at are left
    alone: they become unreachable, and Access reclaims them on its next
    compact.
    """
    dropped = 0
    for base in range(0, len(data) - ACE_PAGE_SIZE + 1, ACE_PAGE_SIZE):
        if data[base] != PAGE_TYPE_DATA:
            continue
        count = int.from_bytes(data[base + 12:base + 14], "little")
        if not count or 14 + 2 * count > ACE_PAGE_SIZE:
            continue
        entries = [int.from_bytes(data[base + 14 + 2 * s:base + 16 + 2 * s],
                                  "little") for s in range(count)]
        starts = sorted({e & 0x0FFF for e in entries})
        for slot, entry in enumerate(entries):
            start = entry & 0x0FFF
            after = [s for s in starts if s > start]
            end = min(after) if after else ACE_PAGE_SIZE
            if _SRP_NAME not in data[base + start:base + end]:
                continue
            data[base + 14 + 2 * slot:base + 16 + 2 * slot] = (
                (entry | _ROW_DELETED) & 0xFFFF).to_bytes(2, "little")
            dropped += 1
    return dropped


# --- LVAL long-value descriptors and chains ----------------------------
#
# A long value is described in MSysAccessStorage by an 8-byte descriptor:
#
#     <u32 length | flags> <u8 slot> <u24 page>
#
# with flags 0x40000000 meaning the pointed row *is* the payload, and
# flags 0 meaning it is the head of a chain whose every row begins with a
# 4-byte <u8 next_slot><u24 next_page> prefix, (0, 0) terminating it.
LVAL_SINGLE_FLAG = 0x40000000


def chain_members(data: bytearray, page: int, slot: int,
                  limit: int = 4096) -> list[tuple[int, int]]:
    """Every ``(page, slot)`` of a chained long value, head first."""
    out: list[tuple[int, int]] = []
    seen: set[tuple[int, int]] = set()
    while (page, slot) not in seen and len(out) < limit:
        seen.add((page, slot))
        out.append((page, slot))
        base = page * ACE_PAGE_SIZE
        start, _ = row_extent(data, base, slot)
        nxt_slot = data[base + start]
        nxt_page = int.from_bytes(data[base + start + 1:base + start + 4],
                                  "little")
        if nxt_page == 0 and nxt_slot == 0:
            return out
        page, slot = nxt_page, nxt_slot
    raise ValueError(f"malformed LVAL chain at page {page} slot {slot}")


def find_lval_descriptor(data: bytearray, page: int, slot: int,
                         current_length: int,
                         storage_page: int | None = None) -> tuple[int, int]:
    """Offset and flags of the descriptor for one long value.

    Located by its exact ``<u32 length|flags><u8 slot><u24 page>`` bytes.
    Eight bytes tying a length to a page and slot is specific enough to
    search the whole file for, which avoids assuming where a database
    keeps its catalog -- ``MSysAccessStorage`` sits on page 48 in a
    freshly created database but need not in general. Pass
    ``storage_page`` to restrict the search to one page.
    """
    pointer = bytes([slot]) + int(page).to_bytes(3, "little")
    if storage_page is None:
        start, stop = 0, len(data)
    else:
        start = storage_page * ACE_PAGE_SIZE
        stop = start + ACE_PAGE_SIZE
    window = bytes(data[start:stop])
    for flags in (LVAL_SINGLE_FLAG, 0):
        needle = (current_length | flags).to_bytes(4, "little") + pointer
        hits = []
        i = window.find(needle)
        while i >= 0:
            hits.append(i)
            i = window.find(needle, i + 1)
        if len(hits) == 1:
            return start + hits[0], flags
        if len(hits) > 1:
            raise ValueError(
                f"long-value descriptor for page {page} slot {slot} "
                f"matched {len(hits)} times")
    raise ValueError(
        f"no long-value descriptor for page {page} slot {slot} "
        f"with length {current_length}")


# Access leaves a 4-byte gap between an LVAL page's slot table and its
# lowest row: a full chain page reads free=4, with the row starting at
# offset 20 rather than 16. Filling that gap is what made a byte-for-byte
# chain rewrite fail to load, so treat it as reserved.
LVAL_ROW_RESERVE = 4


def row_capacity(data: bytearray, page: int, slot: int) -> int:
    """Bytes this row could occupy: its extent plus the page's usable free
    space, keeping :data:`LVAL_ROW_RESERVE` bytes untouched."""
    base = page * ACE_PAGE_SIZE
    start, end = row_extent(data, base, slot)
    slots = slot_offsets(data, base)
    live = [s & 0x0FFF for s in slots if (s & 0xF000) != 0xD000]
    free = min(live) - (14 + 2 * len(slots)) - LVAL_ROW_RESERVE
    return (end - start) + max(0, free)


def tombstone_row(data: bytearray, page: int, slot: int) -> None:
    """Release one LVAL row, marking its slot free and zeroing its bytes.

    The slot keeps its recorded offset and gains the 0xD000 flag, which is
    how Access marks a dead row; the payload is zeroed so no stale bytes
    can be read as part of a neighbouring row, and the page's free-space
    counter is credited.
    """
    base = page * ACE_PAGE_SIZE
    start, end = row_extent(data, base, slot)
    data[base + start:base + end] = bytes(end - start)
    off = base + 14 + 2 * slot
    current = int.from_bytes(data[off:off + 2], "little")
    data[off:off + 2] = (0xD000 | (current & 0x0FFF)).to_bytes(2, "little")
    free = int.from_bytes(data[base + 2:base + 4], "little")
    data[base + 2:base + 4] = (free + (end - start)).to_bytes(2, "little")


def write_chained_lval(data: bytearray, page: int, slot: int,
                       payload: bytes, current_length: int,
                       storage_page: int | None = None) -> None:
    """Rewrite a chained long value across the rows it already occupies.

    The chain keeps its shape -- same rows, same order -- and the payload
    is spread over them in proportion to what each can hold, so no row is
    left empty and no page has to be allocated. The descriptor length is
    updated to match.
    """
    members = chain_members(data, page, slot)
    caps = [row_capacity(data, p, s) - 4 for p, s in members]
    if sum(caps) < len(payload):
        raise ValueError(
            f"chain of {len(members)} rows holds at most {sum(caps)} bytes; "
            f"payload is {len(payload)}")
    # Access fills each chunk to capacity and lets the last one run short
    # (measured: 4072 / 4072 / 583). Match that rather than spreading the
    # payload evenly, which produced a chain Access refused to load.
    shares, left = [], len(payload)
    for cap in caps:
        take = min(cap, left)
        shares.append(take)
        left -= take
    # A payload smaller than the chain needs fewer rows. Access will not
    # load a chain whose trailing chunks carry only their 4-byte link, so
    # the surplus rows are released instead: the chain terminates early
    # and each freed slot is tombstoned.
    used = max(1, sum(1 for share in shares if share))
    pos = 0
    for index in range(used):
        page_no, slot_no = members[index]
        chunk = payload[pos:pos + shares[index]]
        pos += shares[index]
        if index + 1 < used:
            nxt_page, nxt_slot = members[index + 1]
        else:
            nxt_page, nxt_slot = 0, 0
        prefix = bytes([nxt_slot]) + int(nxt_page).to_bytes(3, "little")
        write_row(data, page_no, slot_no, prefix + chunk)
    for page_no, slot_no in members[used:]:
        tombstone_row(data, page_no, slot_no)


def set_lval_payload(data: bytearray, page: int, slot: int, payload: bytes,
                     current_length: int,
                     storage_page: int | None = None) -> None:
    """Replace a long value's bytes, whichever shape it is stored in.

    Single-row values are written in place; chained ones are respread
    over the rows they already occupy. Either way the descriptor's length
    is updated, which is what Access checks on load.
    """
    off, flags = find_lval_descriptor(data, page, slot, current_length,
                                      storage_page)
    if flags & LVAL_SINGLE_FLAG:
        write_row(data, page, slot, payload)
    else:
        write_chained_lval(data, page, slot, payload, current_length,
                           storage_page)
    data[off:off + 4] = (len(payload) | flags).to_bytes(4, "little")
