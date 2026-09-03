"""The module table inside `_VBA_PROJECT`.

After the project cookie comes a module count and then one entry each:

    <u16 stream-name bytes> <stream name UTF-16>
    <u16 20> <10-character cookie UTF-16>
    ff ff <u16 name operand> <u16 name bytes> <name UTF-16>
    ff ff <u16 module cookie> 00*6 <u16 reserve> 00 00 00 <u32 module offset>

Entries are separated by `ff ff`; the first follows the count directly.
After the last one comes the project's own trailer,

    ff ff <u16> 00 00 01 01 <u16 next reserve> 00 00

and the two numbers there are the thing to get right when adding a
module: **a new entry's `reserve` is the trailer's current `next
reserve`, and the trailer then advances by 0x20**.  Measured across four
projects:

| modules | per-module reserve | trailer |
|---|---|---|
| 1 | 0x208 | 0x260 |
| 2 | 0x208, 0x278 | 0x298 |
| 3 | 0x208, 0x278, 0x298 | 0x2b8 |

so Access's own third module took 0x298, exactly the trailer it found.
"""

from __future__ import annotations

from dataclasses import dataclass

SEPARATOR = b"\xff\xff"
COOKIE_BYTES = 20
TAIL = 19  # ff ff plus the 17 fixed bytes that end an entry
RESERVE_STEP = 0x20
TRAILER = bytes.fromhex("00000101")


@dataclass(frozen=True)
class ModuleEntry:
    start: int
    end: int
    stream_name: str
    name: str
    operand: int
    cookie: bytes
    reserve: int
    offset: int


def module_count(blob: bytes, cookie: bytes) -> int:
    """Where the module count sits, just past the project cookie."""
    at = blob.find(cookie)
    if at < 0:
        raise LookupError("no project cookie in _VBA_PROJECT")
    return at + len(cookie)


def entries(blob: bytes, cookie: bytes) -> list[ModuleEntry]:
    """Every module entry, in the order the table holds them."""
    at = module_count(blob, cookie) + 2
    out: list[ModuleEntry] = []
    while True:
        length = int.from_bytes(blob[at : at + 2], "little")
        if length == 0 or length > 512 or at + 2 + length > len(blob):
            return out
        stream_name = blob[at + 2 : at + 2 + length].decode("utf-16-le")
        where = at + 2 + length
        if int.from_bytes(blob[where : where + 2], "little") != COOKIE_BYTES:
            return out
        module_cookie = blob[where + 2 : where + 2 + COOKIE_BYTES]
        where += 2 + COOKIE_BYTES
        if blob[where : where + 2] != SEPARATOR:
            return out
        operand = int.from_bytes(blob[where + 2 : where + 4], "little")
        name_bytes = int.from_bytes(blob[where + 4 : where + 6], "little")
        name = blob[where + 6 : where + 6 + name_bytes].decode("utf-16-le")
        end = where + 6 + name_bytes + TAIL
        tail = blob[where + 6 + name_bytes : end]
        out.append(
            ModuleEntry(
                at,
                end,
                stream_name,
                name,
                operand,
                module_cookie,
                int.from_bytes(tail[10:12], "little"),
                int.from_bytes(tail[15:19], "little"),
            )
        )
        at = end + len(SEPARATOR)


def trailer(blob: bytes, cookie: bytes) -> int:
    """Where the project's `next reserve` word sits, past the last entry."""
    last = entries(blob, cookie)[-1]
    at = blob.find(TRAILER, last.end)
    if at < 0:
        raise LookupError("no project trailer after the module table")
    return at + len(TRAILER)


def entry_bytes(stream_name: str, cookie: bytes, operand: int, name: str, module_cookie: bytes,
                reserve: int, offset: int) -> bytes:
    stream_text, text = stream_name.encode("utf-16-le"), name.encode("utf-16-le")
    tail = bytearray(SEPARATOR + module_cookie + bytes(6) + reserve.to_bytes(2, "little") + bytes(3) + offset.to_bytes(4, "little"))
    return (
        len(stream_text).to_bytes(2, "little")
        + stream_text
        + COOKIE_BYTES.to_bytes(2, "little")
        + cookie
        + SEPARATOR
        + operand.to_bytes(2, "little")
        + len(text).to_bytes(2, "little")
        + text
        + bytes(tail)
    )


def add_module(blob: bytes, cookie: bytes, entry: bytes) -> bytes:
    """Append an entry after the last one, count it, and advance the
    project's reserve."""
    last = entries(blob, cookie)[-1]
    out = bytearray(blob[: last.end] + SEPARATOR + entry + blob[last.end :])
    at = module_count(bytes(out), cookie)
    out[at : at + 2] = (int.from_bytes(out[at : at + 2], "little") + 1).to_bytes(2, "little")
    where = trailer(bytes(out), cookie)
    out[where : where + 2] = (int.from_bytes(out[where : where + 2], "little") + RESERVE_STEP).to_bytes(2, "little")
    return bytes(out)


def next_reserve(blob: bytes, cookie: bytes) -> int:
    at = trailer(blob, cookie)
    return int.from_bytes(blob[at : at + 2], "little")


def entry_span(blob: bytes, stream_name: str, name: str) -> tuple[int, int]:
    """Where a module's entry starts and ends, found from its stream name
    (unique per module) and its own name."""
    stream_bytes = stream_name.encode("utf-16-le")
    at = blob.find(stream_bytes)
    if at < 0:
        raise LookupError(f"_VBA_PROJECT has no entry for stream {stream_name!r}")
    text = name.encode("utf-16-le")
    where = blob.find(text, at)
    if where < 0:
        raise LookupError(f"_VBA_PROJECT has no name record for {name!r}")
    return at - 2, where + len(text) + TAIL


def remove_module(blob: bytes, cookie: bytes, stream_name: str, name: str) -> bytes:
    """Drop a module's entry and take one off the count."""
    start, end = entry_span(blob, stream_name, name)
    if blob[start - 2 : start] == SEPARATOR:
        start -= 2  # the separator belongs to the entry being removed
    out = bytearray(blob[:start] + blob[end:])
    at = module_count(bytes(out), cookie)
    out[at : at + 2] = (int.from_bytes(out[at : at + 2], "little") - 1).to_bytes(2, "little")
    return bytes(out)
