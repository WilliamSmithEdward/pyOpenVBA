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

@dataclass(frozen=True)
class Identifier:
    index: int
    name: str
    id_value: int
    type_byte: int
    record_offset: int

def _printable(b: bytes) -> bool:
    return all(0x20 <= c < 0x7F for c in b)

def parse_identifiers(vba_project_stream: bytes) -> list[Identifier]:
    """Parse the ordered identifier table from a ``_VBA_PROJECT`` stream."""
    s = vba_project_stream
    n = len(s)
    cands: list[tuple[int,int,str,int,int]] = []  # (start, end, name, id, type)
    for p in range(n - 6):
        ln = s[p]
        if not (0 < ln < 64):
            continue
        tb = s[p+1]
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
            cands.append((p, end, nm.decode("ascii"),
                          int.from_bytes(s[ne:ne+2],"little"), tb))
            break
    if not cands:
        return []
    # Longest run of records that chain end->start, scanning greedily.
    by_start = {c[0]: c for c in cands}
    best: list[tuple] = []
    for c in cands:
        chain=[c]; cur=c
        while cur[1] in by_start:
            cur = by_start[cur[1]]
            chain.append(cur)
        if len(chain) > len(best):
            best = chain
    return [Identifier(i, c[2], c[3], c[4], c[0]) for i, c in enumerate(best)]

def resolve_name(operand: int, table: list[Identifier]) -> str | None:
    """Map a p-code ``name`` operand to its identifier, or None."""
    delta = operand - NAME_OPERAND_BASE
    if delta < 0 or delta % 2:
        return None
    idx = delta // 2
    return table[idx].name if 0 <= idx < len(table) else None
