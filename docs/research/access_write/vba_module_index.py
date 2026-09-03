"""The module index `_VBA_PROJECT` keeps past the identifier table.

After the identifier table's `02 ff ff 01 01` sentinel comes a u32 size
and then a hash table of six-byte slots:

    <u16 key> <u16 value> <u16 chain>

An empty slot is all `ff`.  A module past the first has an entry whose
key is its **name operand plus one** and whose value is its index in the
module table, alongside six fixed entries a project always carries
(0x0006, 0x0058, 0x020e, 0x021c, 0x021e, 0x0221).

Placement is `slot = (key >> 1) % slots` with linear probing, which fits
every entry of five projects of one to five modules -- 44 entries, no
exceptions.  Since an operand is `2 * slot + 2`, that hash is the
identifier's own slot number plus one, taken modulo the table.

Access grows the table by two slots per module and rehashes it; the
`chain` field then links a subset of the entries in an order this has not
modelled.  Inserting into the table as it stands, without growing it,
leaves every existing entry where it was.
"""

from __future__ import annotations

SENTINEL = bytes.fromhex("02ffff0101")
SLOT = 6
EMPTY = b"\xff" * SLOT


def table(blob: bytes) -> tuple[int, int]:
    """``(offset of the first slot, slot count)``."""
    at = blob.rfind(SENTINEL)
    if at < 0:
        raise LookupError("no identifier-table sentinel in _VBA_PROJECT")
    start = at + len(SENTINEL)
    size = int.from_bytes(blob[start : start + 4], "little")
    if size % SLOT:
        raise LookupError(f"the module index is {size} bytes, not a whole number of slots")
    return start + 4, size // SLOT


def slots(blob: bytes) -> dict[int, tuple[int, int, int]]:
    """Every filled slot, by position."""
    start, count = table(blob)
    out: dict[int, tuple[int, int, int]] = {}
    for i in range(count):
        at = start + i * SLOT
        if blob[at : at + SLOT] == EMPTY:
            continue
        out[i] = tuple(int.from_bytes(blob[at + 2 * k : at + 2 * k + 2], "little") for k in range(3))  # type: ignore[misc]
    return out


def home(key: int, count: int) -> int:
    return (key >> 1) % count


def place(blob: bytes, key: int, value: int, chain: int = 0xFFFF) -> bytes:
    """Add an entry at its hash position, probing forward for room."""
    start, count = table(blob)
    filled = slots(blob)
    if len(filled) >= count:
        raise LookupError("the module index is full")
    at = home(key, count)
    while at in filled:
        at = (at + 1) % count
    record = b"".join(part.to_bytes(2, "little") for part in (key, value, chain))
    where = start + at * SLOT
    return blob[:where] + record + blob[where + SLOT :]


def rebuild(blob: bytes, extra: int, added: tuple[int, int] | None = None) -> bytes:
    """Grow the table by ``extra`` slots, rehash what it holds and add one
    entry.

    A module's entry chains to the slot of the module before it, and the
    first module's chains to itself; the entries that are not modules
    carry ``ffff``.  That rule reproduces Access's own chains exactly on a
    two-module project and on the three-module one it made by adding to
    it.
    """
    start, count = table(blob)
    held = [(key, value) for _at, (key, value, _chain) in sorted(slots(blob).items())]
    if added is not None:
        held.append(added)
    fresh = count + extra
    if len(held) > fresh:
        raise LookupError("more entries than the grown table holds")
    modules = {value: key for key, value in held if value != 0xFFFF and key >= 0x0221}
    grid: list[bytes | None] = [None] * fresh
    where: dict[int, int] = {}
    for key, value in held:
        at = home(key, fresh)
        while grid[at] is not None:
            at = (at + 1) % fresh
        grid[at] = b""  # claimed; filled once every position is known
        where[key] = at
    for key, value in held:
        chain = 0xFFFF
        if key in modules.values():
            before = modules.get(value - 1, key)
            chain = where[before]
        grid[where[key]] = b"".join(part.to_bytes(2, "little") for part in (key, value, chain))
    return (
        blob[: start - 4]
        + (fresh * SLOT).to_bytes(4, "little")
        + b"".join(cell if cell else EMPTY for cell in grid)
        + blob[start + count * SLOT :]
    )
