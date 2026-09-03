"""The module table inside `_VBA_PROJECT`.

After the project cookie comes a module count and then one entry each:

    <u16 stream-name bytes> <stream name UTF-16>
    <u16 20> <10-character cookie UTF-16>
    ff ff <u16 name operand> <u16 name bytes> <name UTF-16>
    ff ff <u16 module cookie> 00*6 <u16> 00 00 00 <u32 module offset>

Entries are separated by `ff ff`; the first follows the count directly.
Measured on a two-module project, where Module1's entry ran 1340..1458
and Alpha's 1461..1575 with the separator between them.
"""

from __future__ import annotations

TAIL = 19  # ff ff plus the 17 fixed bytes that end an entry
COOKIE = 22  # the cookie's length word and its 20 bytes
HEAD = 6  # ff ff plus the operand and the name's length word


def entry_span(blob: bytes, stream_name: str, name: str) -> tuple[int, int]:
    """Where a module's entry starts and ends, found from its stream name
    (unique per module) and its own name."""
    stream_bytes = stream_name.encode("utf-16-le")
    at = blob.find(stream_bytes)
    if at < 0:
        raise LookupError(f"_VBA_PROJECT has no entry for stream {stream_name!r}")
    start = at - 2
    text = name.encode("utf-16-le")
    where = blob.find(text, at)
    if where < 0:
        raise LookupError(f"_VBA_PROJECT has no name record for {name!r}")
    return start, where + len(text) + TAIL


def module_count(blob: bytes, cookie: bytes) -> int:
    at = blob.find(cookie)
    if at < 0:
        raise LookupError("no project cookie in _VBA_PROJECT")
    return at + len(cookie)


def remove_module(blob: bytes, cookie: bytes, stream_name: str, name: str) -> bytes:
    """Drop a module's entry and take one off the count."""
    start, end = entry_span(blob, stream_name, name)
    if blob[start - 2 : start] == b"\xff\xff":
        start -= 2  # the separator belongs to the entry being removed
    out = bytearray(blob[:start] + blob[end:])
    at = module_count(bytes(out), cookie)
    out[at : at + 2] = (int.from_bytes(out[at : at + 2], "little") - 1).to_bytes(2, "little")
    return bytes(out)
