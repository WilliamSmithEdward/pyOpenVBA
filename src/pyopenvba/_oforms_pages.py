"""
A MultiPage's tabs: the TabStrip's parallel arrays and the ``x`` bookkeeping.

A Page is two things at once.  It is a container control, with a site in
its MultiPage's ``f`` and a storage of its own; and it is a *tab*, whose
caption, tip, tag and accelerator live in five parallel arrays inside the
MultiPage's TabStrip record, with one flag word per tab after it and its
position in the ``x`` stream.  Adding or removing one moves all of that
together, which is why it needs a module rather than a branch.

Everything here was measured against Excel adding and removing a page on
the same form ([MS-OFORMS] 2.1.2.3 for the ``x`` layout):

- the five arrays stay the same length as each other and as the tab count;
- ``TabData`` is the tab count, and ``TabsAllocated`` is capacity -- Excel
  keeps it at count + 2 and never shrinks it;
- the TabStrip's tail is one 32-bit flag word per tab;
- ``x`` holds one more PageProperties record than there are pages (the
  first is ignored), then a MultiPageProperties whose tail is the page
  site ids in page order.

Private module: the public surface is :mod:`pyopenvba.forms`.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from pyopenvba.exceptions import FormParseError

# MultiPagePropMask ([MS-OFORMS] 2.2.4): PageCount, then ID.
_MASK_PAGE_COUNT = 1 << 1
_MASK_ID = 1 << 2

_RECORD_VERSION = (0x00, 0x02)

# An empty PageProperties record: header, zero mask, no data.
EMPTY_PAGE_PROPERTIES = struct.pack("<BBHI", 0x00, 0x02, 0x0004, 0)

# What Excel writes for a tab it has just created.
NEW_TAB_FLAGS = 0x00000003
# Capacity Excel keeps ahead of the tab count.
TAB_HEADROOM = 2


@dataclass
class TabString:
    """One entry of a TabStrip array, kept decoded and as stored bytes."""

    text: str
    compressed: bool
    raw: bytes = b""
    pad: bytes = b""
    """Alignment bytes after the string; the spec leaves them undefined,
    so they are replayed rather than recomputed as zeros."""


def parse_string_array(blob: bytes, encoding: str) -> list[TabString]:
    """Split one TabStrip array into its entries.

    Each is a ``CountOfBytesWithCompressionFlag`` and its bytes, padded to
    a 4-byte boundary.
    """
    entries: list[TabString] = []
    pos = 0
    while pos + 4 <= len(blob):
        packed = int(struct.unpack_from("<I", blob, pos)[0])
        pos += 4
        length = packed & 0x7FFFFFFF
        compressed = bool(packed & 0x80000000)
        if pos + length > len(blob):
            raise FormParseError("TabStrip array entry runs past the end")
        raw = blob[pos:pos + length]
        pos += length
        over = pos % 4
        pad = blob[pos:pos + (4 - over)] if over else b""
        pos += len(pad)
        entries.append(
            TabString(
                raw.decode(encoding if compressed else "utf-16-le", "replace"),
                compressed,
                raw,
                pad,
            )
        )
    return entries


def serialize_string_array(entries: list[TabString], encoding: str) -> bytes:
    """Rebuild a TabStrip array.  Unedited entries round-trip exactly."""
    out = bytearray()
    for entry in entries:
        raw = entry.raw
        if raw == b"" and entry.text:
            raw = entry.text.encode(
                encoding if entry.compressed else "utf-16-le", "replace"
            )
        packed = (len(raw) & 0x7FFFFFFF) | (0x80000000 if entry.compressed else 0)
        out += struct.pack("<I", packed) + raw
        over = len(out) % 4
        if over:
            gap = 4 - over
            out += entry.pad if len(entry.pad) == gap else bytes(gap)
    return bytes(out)


def new_tab_string(text: str) -> TabString:
    """A fresh entry, compressed when every character fits one byte."""
    return TabString(text, all(ord(c) <= 0xFF for c in text))


@dataclass
class PageBookkeeping:
    """The ``x`` stream: which pages a MultiPage has, and in what order."""

    page_props: list[bytes] = field(default_factory=lambda: [])
    """Raw PageProperties records.  There is one more than there are
    pages and the first is ignored, so these are carried rather than
    interpreted."""
    mask: int = _MASK_PAGE_COUNT | _MASK_ID
    identifier: int = 0
    page_ids: list[int] = field(default_factory=lambda: [])

    def add(self, page_id: int) -> None:
        self.page_ids.append(page_id)
        self.page_props.append(EMPTY_PAGE_PROPERTIES)

    def remove(self, page_id: int) -> None:
        if page_id not in self.page_ids:
            raise FormParseError(f"page id {page_id} is not in the page bookkeeping")
        self.page_ids.remove(page_id)
        self.page_props.pop()


def parse_page_bookkeeping(blob: bytes) -> PageBookkeeping:
    """Read the ``x`` stream, refusing anything that does not reconcile."""
    book = PageBookkeeping()
    pos = 0
    while pos < len(blob):
        start = pos
        if pos + 4 > len(blob):
            break
        minor, major, cb = struct.unpack_from("<BBH", blob, pos)
        if (minor, major) != _RECORD_VERSION:
            raise FormParseError(f"x stream: not a record at {start}")
        pos += 4
        if pos + 4 > len(blob):
            raise FormParseError("x stream: record has no mask")
        mask = int(struct.unpack_from("<I", blob, pos)[0])
        record_end = start + 4 + cb
        if record_end > len(blob):
            raise FormParseError("x stream: record runs past the end")
        if not mask & _MASK_PAGE_COUNT:
            # A PageProperties record: carried whole.
            book.page_props.append(blob[start:record_end])
            pos = record_end
            continue
        # The MultiPageProperties: its tail is the PageIDs array.
        page_count = int(struct.unpack_from("<i", blob, pos + 4)[0])
        book.mask = mask
        if mask & _MASK_ID:
            book.identifier = int(struct.unpack_from("<i", blob, pos + 8)[0])
        pos = record_end
        while pos + 4 <= len(blob):
            book.page_ids.append(int(struct.unpack_from("<i", blob, pos)[0]))
            pos += 4
        if pos != len(blob):
            raise FormParseError("x stream: trailing bytes after PageIDs")
        if len(book.page_ids) != page_count:
            raise FormParseError(
                f"x stream: {len(book.page_ids)} PageIDs for PageCount {page_count}"
            )
        return book
    raise FormParseError("x stream has no MultiPageProperties")


def serialize_page_bookkeeping(book: PageBookkeeping) -> bytes:
    """Rebuild the ``x`` stream.  Unedited bookkeeping round-trips exactly."""
    out = bytearray(b"".join(book.page_props))
    cb = 4 + 4 + (4 if book.mask & _MASK_ID else 0)
    out += struct.pack("<BBHIi", 0x00, 0x02, cb, book.mask, len(book.page_ids))
    if book.mask & _MASK_ID:
        out += struct.pack("<i", book.identifier)
    for page_id in book.page_ids:
        out += struct.pack("<i", page_id)
    return bytes(out)
