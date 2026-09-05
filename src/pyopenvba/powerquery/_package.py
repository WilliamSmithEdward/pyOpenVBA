"""The OPC package inside a DataMashup blob.

Three parts live in it -- ``Config/Package.xml``, ``[Content_Types].xml``
and ``Formulas/Section1.m`` -- held in a ZIP that Excel writes with
settings of its own: raw deflate at level 6, a growth-hint extra field on
every local header, and version words 45 and 20.  All of that is
measured, so :meth:`Package.serialize` reproduces Excel's bytes for the
same parts rather than merely producing a ZIP that opens.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field

from pyopenvba._deflate import raw_compress
from pyopenvba.exceptions import PowerQueryError

#: The parts Excel writes, in the order it writes them.
CONFIG_PART = "Config/Package.xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
SECTION_PART = "Formulas/Section1.m"
PART_ORDER = (CONFIG_PART, CONTENT_TYPES_PART, SECTION_PART)

#: The version words Excel puts in the headers.
MADE_BY = 45
NEEDED = 20
#: Bits 1 and 2 of the general-purpose flag, which the ZIP format reads
#: as "maximum compression".
FLAGS = 0x0002
DEFLATED = 8

#: Every local header carries the Open Packaging growth hint, and every
#: central header carries none.  These are the exact 28 bytes Excel
#: writes: the 0xA220 tag, a 24-byte body of the 0xA028 signature, a
#: 20-byte hint and its padding.
GROWTH_HINT = bytes.fromhex("20a21800") + bytes.fromhex("28a01400") + bytes(20)

#: The two documents that do not change from workbook to workbook.
DEFAULT_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="text/xml" />'
    '<Default Extension="m" ContentType="application/x-ms-m" />'
    "</Types>"
)


def default_config(version: str = "2.157.151.0", min_version: str = "2.21.0.0", culture: str = "en-US") -> str:
    """``Config/Package.xml`` as Excel writes it."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<Package xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<Version>{version}</Version><MinVersion>{min_version}</MinVersion>"
        f"<Culture>{culture}</Culture></Package>"
    )


@dataclass
class Part:
    """One entry of the package."""

    name: str
    data: bytes
    #: The DOS date and time the entry carries, kept so a rewrite of one
    #: part leaves the others exactly as they were.
    dos_time: int = 0
    dos_date: int = 0x21 << 5 | 1  # 1980-01-01, what a zeroed stamp means


@dataclass
class Package:
    """The package as parts, with the bytes it came from."""

    parts: list[Part] = field(default_factory=lambda: [])
    #: The bytes read, kept so an untouched package writes back exactly.
    source: bytes | None = None

    # -- reading ------------------------------------------------------------

    @classmethod
    def parse(cls, raw: bytes) -> Package:
        parts: list[Part] = []
        at = 0
        while raw[at : at + 4] == b"PK\x03\x04":
            (_need, _flags, method, dos_time, dos_date, _crc, csize, usize, nlen, xlen) = struct.unpack_from(
                "<HHHHHIIIHH", raw, at + 4
            )
            name = raw[at + 30 : at + 30 + nlen].decode("utf-8")
            body_at = at + 30 + nlen + xlen
            body = raw[body_at : body_at + csize]
            if method == 0:
                data = body
            elif method == DEFLATED:
                try:
                    data = zlib.decompress(body, -15)
                except zlib.error as exc:
                    raise PowerQueryError(f"the package part {name!r} does not inflate: {exc}") from exc
            else:
                raise PowerQueryError(f"the package part {name!r} uses compression method {method}")
            if len(data) != usize:
                raise PowerQueryError(f"the package part {name!r} is {len(data)} bytes, not {usize}")
            parts.append(Part(name, data, dos_time, dos_date))
            at = body_at + csize
        if not parts:
            raise PowerQueryError("this package holds no parts")
        return cls(parts=parts, source=raw)

    def part(self, name: str) -> Part:
        for part in self.parts:
            if part.name == name:
                return part
        raise PowerQueryError(f"this package has no part named {name!r}")

    def has(self, name: str) -> bool:
        return any(part.name == name for part in self.parts)

    def read(self, name: str) -> bytes:
        return self.part(name).data

    # -- writing ------------------------------------------------------------

    def write(self, name: str, data: bytes) -> None:
        """Replace a part's bytes, or add the part at the end."""
        for part in self.parts:
            if part.name == name:
                if part.data != data:
                    part.data = data
                    self.source = None
                return
        stamp = self.parts[-1] if self.parts else Part(name, b"")
        self.parts.append(Part(name, data, stamp.dos_time, stamp.dos_date))
        self.source = None

    def serialize(self) -> bytes:
        """The package's bytes: the ones it was read from when nothing has
        changed, and Excel's own layout otherwise."""
        if self.source is not None:
            return self.source
        out = bytearray()
        central = bytearray()
        for part in self.parts:
            name = part.name.encode("utf-8")
            body = raw_compress(part.data)
            crc = zlib.crc32(part.data) & 0xFFFFFFFF
            offset = len(out)
            out += b"PK\x03\x04" + struct.pack(
                "<HHHHHIIIHH", NEEDED, FLAGS, DEFLATED, part.dos_time, part.dos_date,
                crc, len(body), len(part.data), len(name), len(GROWTH_HINT),
            )
            out += name + GROWTH_HINT + body
            central += b"PK\x01\x02" + struct.pack(
                "<HHHHHHIIIHHHHHII", MADE_BY, NEEDED, FLAGS, DEFLATED, part.dos_time, part.dos_date,
                crc, len(body), len(part.data), len(name), 0, 0, 0, 0, 0, offset,
            )
            central += name
        start = len(out)
        out += central
        out += b"PK\x05\x06" + struct.pack(
            "<HHHHIIH", 0, 0, len(self.parts), len(self.parts), len(central), start, 0
        )
        return bytes(out)


def new_package(section: str) -> Package:
    """A package holding a section document, laid out as Excel lays it."""
    package = Package()
    package.parts = [
        Part(CONFIG_PART, default_config().encode("utf-8")),
        Part(CONTENT_TYPES_PART, DEFAULT_CONTENT_TYPES.encode("utf-8")),
        Part(SECTION_PART, section.encode("utf-8")),
    ]
    return package
