"""Identifier-table parsing and p-code name resolution (host-agnostic).

The identifier table lives in the project's ``_VBA_PROJECT`` stream
(magic ``CC 61``), which is byte-compatible across Excel / Word /
PowerPoint / Access. Compiled p-code ``name`` operands do NOT carry the
table's ``id`` value; they are ``NAME_OPERAND_BASE + 2*index`` where
``index`` is the record's ordinal position in that table.

Record layout::

    <u8 name_len> <u8 type> [<descriptor bytes>] <ASCII name>
    <u16 id> <0x10> <0x00>

Some type bytes insert descriptor bytes between the type and the name,
so records are anchored on the trailing ``10 00`` marker rather than
chained forward from a guessed start.
"""
from __future__ import annotations

from dataclasses import dataclass

NAME_OPERAND_BASE = 0x20E   # operand = BASE + 2*index (empirically fixed)

# Operands BELOW NAME_OPERAND_BASE address a second identifier space: a
# fixed table of built-in names held by the VBA runtime, not stored in
# the file at all. It is identical for every project and ordered
# alphabetically (case-insensitively), so operands can be mapped by
# probing and the result shipped as a lookup.
#
# A user identifier that collides with a built-in name reuses the
# built-in's operand instead of gaining a project-table entry, which is
# why e.g. `Dim b` and `Dim f` never appear in the project table.
#
# Partial map, each entry confirmed by compiling a probe with Excel and
# reading back the operand. Extend by probing more names; the ordering
# invariant (sorted by operand == sorted alphabetically) is a useful
# self-check on any addition.
BUILTIN_OPERANDS: dict[int, str] = {
    0x0012: "Array",
    0x0018: "b",
    0x0034: "CDec",
    0x003A: "ChDir",
    0x004C: "CurDir",
    0x005A: "Date",
    0x007E: "Dir",
    0x0084: "DoEvents",
    0x009A: "Error",
    0x00A4: "f",
    0x00AC: "Format",
    0x00B0: "FreeFile",
    0x00C8: "Input",
    0x00DC: "Left",
    0x00FA: "Mid",
    0x00FE: "MidB",
    0x0134: "Randomize",
    0x0140: "RGB",
    0x0146: "Seek",
    0x015C: "String",
}

@dataclass(frozen=True)
class Identifier:
    index: int
    name: str
    id_value: int
    type_byte: int
    record_offset: int

def _printable(b: bytes) -> bool:
    """True for bytes that could spell an identifier.

    Identifiers are stored MBCS-encoded in the project code page, so a
    name may legitimately contain bytes above 0x7F (accented Latin,
    Cyrillic, CJK). Restricting this to ASCII silently drops those
    records -- which hid every non-ASCII identifier until the hash work
    surfaced it. Only C0 controls and DEL are rejected.
    """
    return all(c >= 0x20 and c != 0x7F for c in b)

def _plausible_type(tb: int) -> bool:
    """Every observed record type is a multiple of 4, at most 0xAC.

    Without this check a byte that happens to be ASCII (e.g. 0x45 'E')
    can pose as a type and produce a false record one byte before the
    real table, yielding a chain of equal length that starts
    mid-record (symptom: a leading 'xcel' / 'cel' instead of 'Excel').
    """
    return tb % 4 == 0 and tb <= 0xAC

def parse_identifiers(vba_project_stream: bytes,
                      code_page: str = "cp1252") -> list[Identifier]:
    """Parse the ordered identifier table from a ``_VBA_PROJECT`` stream.

    Names are stored MBCS-encoded in the project's code page (its
    ``PROJECTCODEPAGE`` record), so pass that code page to read
    identifiers outside ASCII correctly.
    """
    s = vba_project_stream
    n = len(s)
    cands: list[tuple[int,int,str,int,int,int]] = []  # (start,end,name,id,type,desc)
    for p in range(n - 6):
        ln = s[p]
        if not (0 < ln < 64):
            continue
        tb = s[p+1]
        if not _plausible_type(tb):
            continue
        for desc in (0, 6):           # some types carry a 6-byte descriptor
            ns = p + 2 + desc
            ne = ns + ln
            end = ne + 4
            if end > n:
                continue
            if s[end-2] != 0x10 or s[end-1] != 0x00:
                continue
            nm = s[ns:ne]
            if not _printable(nm):
                continue
            cands.append((p, end, nm.decode(code_page, errors="replace"),
                          int.from_bytes(s[ne:ne+2],"little"), tb, desc))
    if not cands:
        return []
    # Longest run of records that chain end->start.
    #
    # Tie-break matters: the 6-byte-descriptor variant can produce a
    # false record that coincidentally aligns a few bytes before the
    # real table start and yields a chain of equal length beginning
    # mid-record (observed: a bogus leading 'cel' instead of 'Excel').
    # Among equal-length chains prefer the one whose first record
    # needs no descriptor, then the latest start -- the real table
    # begins at a plain record.
    by_start: dict[int, tuple] = {}
    for c in cands:
        # index by start; prefer the descriptor-free interpretation
        if c[0] not in by_start or c[5] < by_start[c[0]][5]:
            by_start[c[0]] = c
    best: list[tuple] = []
    best_key = None
    for c in cands:
        chain=[c]
        cur=c
        while cur[1] in by_start:
            cur = by_start[cur[1]]
            chain.append(cur)
        key = (len(chain), c[5] == 0, c[0])
        if best_key is None or key > best_key:
            best_key = key
            best = chain
    return [Identifier(i, c[2], c[3], c[4], c[0]) for i, c in enumerate(best)]

def resolve_name(operand: int, table: list[Identifier]) -> str | None:
    """Map a p-code ``name`` operand to its identifier, or None.

    Operands at or above :data:`NAME_OPERAND_BASE` index the project
    table; lower operands address the runtime's built-in table (see
    :data:`BUILTIN_OPERANDS`).
    """
    if operand < NAME_OPERAND_BASE:
        return BUILTIN_OPERANDS.get(operand)
    delta = operand - NAME_OPERAND_BASE
    if delta % 2:
        return None
    idx = delta // 2
    return table[idx].name if 0 <= idx < len(table) else None


def builtin_table_is_ordered() -> bool:
    """Self-check: the built-in map must stay alphabetically ordered."""
    names = [BUILTIN_OPERANDS[k] for k in sorted(BUILTIN_OPERANDS)]
    return names == sorted(names, key=str.lower)
