"""Just enough of the OPC container to change one part and leave the rest.

A workbook is a ZIP, and the Power Query package lives in one part of
it.  Rewriting the file with a general-purpose ZIP writer would re-deflate
every other part and rewrite every header, so a one-query edit would
change megabytes for no reason.  This reader keeps each entry's bytes as
they arrived -- header, growth-hint padding and compressed body alike --
and writes them back untouched, which leaves an unchanged workbook
identical to the byte and a changed one different only where the change
is.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

from pyopenvba._deflate import raw_compress
from pyopenvba.exceptions import PowerQueryError

_LOCAL = b"PK\x03\x04"
_CENTRAL = b"PK\x01\x02"
_END = b"PK\x05\x06"
_END64_LOCATOR = b"PK\x06\x07"
_DEFLATED = 8
_STORED = 0
#: What Excel writes on the parts of a workbook.
_FLAGS = 0x0006
_MADE_BY = 45
_NEEDED = 20


@dataclass
class Entry:
    """One part, with everything needed to write it back as it was."""

    name: str
    #: The compressed bytes exactly as they were stored.
    body: bytes
    method: int
    flags: int
    dos_time: int
    dos_date: int
    crc: int
    uncompressed_size: int
    local_extra: bytes = b""
    central_extra: bytes = b""
    comment: bytes = b""
    made_by: int = _MADE_BY
    needed: int = _NEEDED
    disk: int = 0
    internal_attributes: int = 0
    external_attributes: int = 0

    def read(self) -> bytes:
        if self.method == _STORED:
            return self.body
        if self.method != _DEFLATED:
            raise PowerQueryError(f"the part {self.name!r} uses compression method {self.method}")
        try:
            return zlib.decompress(self.body, -15)
        except zlib.error as exc:
            raise PowerQueryError(f"the part {self.name!r} does not inflate: {exc}") from exc


@dataclass
class OpcFile:
    """A package read from bytes, and written back from them."""

    entries: list[Entry] = field(default_factory=lambda: [])
    source: bytes | None = field(default=None, repr=False)

    # -- reading ------------------------------------------------------------

    @classmethod
    def parse(cls, raw: bytes) -> OpcFile:
        end = raw.rfind(_END)
        if end < 0:
            raise PowerQueryError("this file has no ZIP end record; it is not an Office package")
        if raw.rfind(_END64_LOCATOR) >= 0:
            raise PowerQueryError("ZIP64 packages are not handled here")
        count, _size, offset = struct.unpack_from("<HII", raw, end + 10)
        entries: list[Entry] = []
        at = offset
        for _ in range(count):
            if raw[at : at + 4] != _CENTRAL:
                raise PowerQueryError(f"the central directory breaks off at {at}")
            (
                made_by, needed, flags, method, dos_time, dos_date, crc, csize, usize,
                name_length, extra_length, comment_length, disk, internal, external, local_at,
            ) = struct.unpack_from("<HHHHHHIIIHHHHHII", raw, at + 4)
            name = raw[at + 46 : at + 46 + name_length].decode("utf-8")
            central_extra = raw[at + 46 + name_length : at + 46 + name_length + extra_length]
            comment_at = at + 46 + name_length + extra_length
            comment = raw[comment_at : comment_at + comment_length]
            if raw[local_at : local_at + 4] != _LOCAL:
                raise PowerQueryError(f"the part {name!r} has no local header at {local_at}")
            local_name_length, local_extra_length = struct.unpack_from("<HH", raw, local_at + 26)
            body_at = local_at + 30 + local_name_length + local_extra_length
            entries.append(
                Entry(
                    name=name,
                    body=raw[body_at : body_at + csize],
                    method=method,
                    flags=flags,
                    dos_time=dos_time,
                    dos_date=dos_date,
                    crc=crc,
                    uncompressed_size=usize,
                    local_extra=raw[local_at + 30 + local_name_length : body_at],
                    central_extra=central_extra,
                    comment=comment,
                    made_by=made_by,
                    needed=needed,
                    disk=disk,
                    internal_attributes=internal,
                    external_attributes=external,
                )
            )
            at = comment_at + comment_length
        return cls(entries=entries, source=raw)

    def names(self) -> list[str]:
        return [entry.name for entry in self.entries]

    def has(self, name: str) -> bool:
        return any(entry.name == name for entry in self.entries)

    def entry(self, name: str) -> Entry:
        for entry in self.entries:
            if entry.name == name:
                return entry
        raise PowerQueryError(f"this package has no part named {name!r}")

    def read(self, name: str) -> bytes:
        return self.entry(name).read()

    # -- writing ------------------------------------------------------------

    def write(self, name: str, data: bytes, *, after: str | None = None) -> None:
        """Replace a part, or add one.  A new part goes after `after` when
        that part is there, which is where Excel keeps a customXml item:
        beside the ones already in the package."""
        body = raw_compress(data)
        fresh = Entry(
            name=name,
            body=body,
            method=_DEFLATED,
            flags=_FLAGS,
            dos_time=0,
            dos_date=0x21,
            crc=zlib.crc32(data) & 0xFFFFFFFF,
            uncompressed_size=len(data),
        )
        for index, entry in enumerate(self.entries):
            if entry.name == name:
                if entry.read() == data:
                    return
                fresh.dos_time, fresh.dos_date = entry.dos_time, entry.dos_date
                fresh.flags = entry.flags
                self.entries[index] = fresh
                self.source = None
                return
        at = len(self.entries)
        if after is not None and self.has(after):
            at = self.entries.index(self.entry(after)) + 1
        self.entries.insert(at, fresh)
        self.source = None

    def remove(self, name: str) -> None:
        kept = [entry for entry in self.entries if entry.name != name]
        if len(kept) != len(self.entries):
            self.entries = kept
            self.source = None

    def serialize(self) -> bytes:
        if self.source is not None:
            return self.source
        out = bytearray()
        central = bytearray()
        for entry in self.entries:
            name = entry.name.encode("utf-8")
            offset = len(out)
            out += _LOCAL + struct.pack(
                "<HHHHHIIIHH", entry.needed, entry.flags, entry.method, entry.dos_time,
                entry.dos_date, entry.crc, len(entry.body), entry.uncompressed_size,
                len(name), len(entry.local_extra),
            )
            out += name + entry.local_extra + entry.body
            central += _CENTRAL + struct.pack(
                "<HHHHHHIIIHHHHHII", entry.made_by, entry.needed, entry.flags, entry.method,
                entry.dos_time, entry.dos_date, entry.crc, len(entry.body),
                entry.uncompressed_size, len(name), len(entry.central_extra),
                len(entry.comment), entry.disk, entry.internal_attributes,
                entry.external_attributes, offset,
            )
            central += name + entry.central_extra + entry.comment
        start = len(out)
        out += central
        out += _END + struct.pack(
            "<HHHHIIH", 0, 0, len(self.entries), len(self.entries), len(central), start, 0
        )
        return bytes(out)
