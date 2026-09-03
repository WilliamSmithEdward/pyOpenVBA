"""A module's name inside `_VBA_PROJECT`, renamed the way Access renames it.

Measured by diffing Access's own rename of `Alpha` to `Gamma` in a project
that also had a module calling `Alpha.AlphaGo`:

* The identifier table keeps the old name and **appends** a new record,
  `<u8 len> <u8 kind=4> <name> <u16 hash> 10 00`, immediately before the
  `02 ff ff 01 01` sentinel.  The hash is the OLE `LHashValOfNameSysA`
  value the p-code research solved (`Gamma` -> 0xc385, written `85 c3`).
* Two u16 counters before the table move: a slot counter at
  ``start - 14`` and a record counter at ``start - 12``, each by one.
  The records themselves begin after a `00 00 00 00 02 00` anchor.
* The name a module binds to is a UTF-16 record,
  ``ff ff <u16 operand> <u16 byte length> <name UTF-16> ff ff``, and the
  operand is **2 * slot + 2** -- the slot the appended record took, which
  is the counter's value before it was bumped.  Access wrote 910 for a
  slot counter of 454.

Appending rather than renaming in place is what keeps a variable that
happens to share the module's name pointing at its own record: VBA is
case-insensitive and one record serves every use of a name.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from pcode_hash import identifier_hash  # noqa: E402

ANCHOR = bytes.fromhex("000000000200")
SENTINEL = bytes.fromhex("02ffff0101")
SEPARATOR = bytes.fromhex("1000")
KIND_NAME = 0x04
SLOT_COUNTER = 14  # bytes before the first record
RECORD_COUNTER = 12


def table_start(blob: bytes) -> int:
    """The anchor also occurs in unrelated data, so a candidate counts
    only when the record behind it reads as a named one."""
    at = -1
    while True:
        at = blob.find(ANCHOR, at + 1)
        if at < 0:
            raise LookupError("no identifier-table anchor in _VBA_PROJECT")
        start = at + len(ANCHOR)
        length = blob[start]
        name = blob[start + 2 : start + 2 + length]
        if 0 < length < 64 and name.isascii() and name.decode("latin-1").isprintable():
            return start


def identifier_record(name: str) -> bytes:
    text = name.encode("latin-1")
    return bytes((len(text), KIND_NAME)) + text + identifier_hash(name).to_bytes(2, "little") + SEPARATOR


def append_identifier(blob: bytes, name: str) -> tuple[bytes, int]:
    """Add a name to the project identifier table and return the blob with
    the p-code operand the new record answers to."""
    start = table_start(blob)
    stop = blob.rfind(SENTINEL)
    if stop < start:
        raise LookupError("the identifier table has no sentinel after its start")
    slot = int.from_bytes(blob[start - SLOT_COUNTER : start - SLOT_COUNTER + 2], "little")
    out = bytearray(blob[:stop] + identifier_record(name) + blob[stop:])
    for back in (SLOT_COUNTER, RECORD_COUNTER):
        at = start - back
        out[at : at + 2] = (int.from_bytes(out[at : at + 2], "little") + 1).to_bytes(2, "little")
    return bytes(out), 2 * slot + 2


def module_record(blob: bytes, name: str) -> int:
    """Where the `ff ff <operand> <length> <name UTF-16> ff ff` record for
    this module starts, at its operand field."""
    text = name.encode("utf-16-le")
    at = 0
    while True:
        at = blob.find(text, at)
        if at < 0:
            raise LookupError(f"_VBA_PROJECT has no UTF-16 record for {name!r}")
        if int.from_bytes(blob[at - 2 : at], "little") == len(text) and blob[at - 6 : at - 4] == b"\xff\xff":
            return at - 4
        at += len(text)


def rename_in_vba_project(blob: bytes, old: str, new: str) -> tuple[bytes, str]:
    """Append the new name and point the module's record at it."""
    at = module_record(blob, old)
    was = int.from_bytes(blob[at : at + 2], "little")
    blob, operand = append_identifier(blob, new)
    out = bytearray(blob)
    text = new.encode("utf-16-le")
    out[at : at + 4 + len(old.encode("utf-16-le"))] = (
        operand.to_bytes(2, "little") + len(text).to_bytes(2, "little") + text
    )
    return bytes(out), f"_VBA_PROJECT (identifier appended, operand {was} -> {operand})"


# --- the per-module flag list ------------------------------------------------
# Ahead of the module table sits `<u16 module count> <count u16 flags, each
# 1> <u16 n> <n records of four bytes whose first word ascends by two>`.
# Adding a module bumps the count and inserts one more flag; leaving it
# alone gives a project Access lists through its own catalog but the VBE
# calls corrupt. Measured at the same offset in four projects, with the
# count following the module table's own count exactly.
FLAG = (1).to_bytes(2, "little")


def flag_list(blob: bytes, modules: int, window: range = range(1000, 1400, 2)) -> int:
    """Where the flag list's count sits."""
    for at in window:
        if int.from_bytes(blob[at : at + 2], "little") != modules:
            continue
        if any(blob[at + 2 + 2 * i : at + 4 + 2 * i] != FLAG for i in range(modules)):
            continue
        after = at + 2 + 2 * modules
        count = int.from_bytes(blob[after : after + 2], "little")
        if not 1 <= count <= 64:
            continue
        operands = [int.from_bytes(blob[after + 2 + 4 * i : after + 4 + 4 * i], "little") for i in range(count)]
        if len(operands) > 1 and all(operands[i + 1] == operands[i] + 2 for i in range(len(operands) - 1)):
            return at
    raise LookupError("no per-module flag list in _VBA_PROJECT")


def add_module_flag(blob: bytes, modules: int) -> bytes:
    """Count one more module and give it its flag."""
    at = flag_list(blob, modules)
    out = bytearray(blob)
    out[at : at + 2] = (modules + 1).to_bytes(2, "little")
    where = at + 2 + 2 * modules
    return bytes(out[:where] + bytearray(FLAG) + out[where:])
