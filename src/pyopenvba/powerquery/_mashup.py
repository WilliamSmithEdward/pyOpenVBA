"""The DataMashup blob: the envelope Excel keeps a Power Query package in.

Five pieces sit one after another, each but the first behind its own
length::

    u32 version
    u32 length + the OPC package
    u32 length + the permission list, as XML
    u32 length + the metadata section
    u32 length + the permission bindings

What each piece is for was settled against live Excel rather than
assumed.  The permission list and the metadata's content package are
both load-bearing: a workbook whose permissions are empty, or whose
metadata is cut short of its content, opens with an error instead of its
queries.  The bindings are not: they hold a signature protected by the
Windows data-protection API, tied to the machine that wrote them, and
Excel opens and refreshes a workbook whose bindings are empty.  So a
package written here drops them rather than carrying a signature that no
longer covers what it signs.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._metadata import Metadata
from pyopenvba.powerquery._package import Package

#: The only version this format has been seen with.
MASHUP_VERSION = 0

#: The permission list Excel writes for a workbook whose queries carry no
#: credentials of their own.
DEFAULT_PERMISSIONS = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    "<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>"
    "<FirewallEnabled>true</FirewallEnabled></PermissionList>"
).encode("utf-8")


def _section(raw: bytes, at: int, what: str) -> tuple[bytes, int]:
    if at + 4 > len(raw):
        raise PowerQueryError(f"the mashup ends where its {what} length should be")
    length = struct.unpack_from("<I", raw, at)[0]
    end = at + 4 + length
    if end > len(raw):
        raise PowerQueryError(f"the mashup {what} claims {length} bytes and the blob is shorter")
    return raw[at + 4 : end], end


@dataclass
class Mashup:
    """A parsed DataMashup blob."""

    package: Package
    metadata: Metadata
    permissions: bytes = DEFAULT_PERMISSIONS
    bindings: bytes = b""
    version: int = MASHUP_VERSION
    #: The bytes this was read from, kept so an untouched blob writes back
    #: exactly as it arrived.
    source: bytes | None = field(default=None, repr=False)

    @classmethod
    def parse(cls, raw: bytes) -> Mashup:
        if len(raw) < 8:
            raise PowerQueryError("a mashup blob is at least eight bytes")
        version = struct.unpack_from("<I", raw, 0)[0]
        package_bytes, at = _section(raw, 4, "package")
        permissions, at = _section(raw, at, "permission list")
        metadata_bytes, at = _section(raw, at, "metadata")
        bindings, at = _section(raw, at, "permission bindings")
        if at != len(raw):
            raise PowerQueryError(f"the mashup carries {len(raw) - at} bytes past its last section")
        return cls(
            package=Package.parse(package_bytes),
            metadata=Metadata.parse(metadata_bytes),
            permissions=permissions,
            bindings=bindings,
            version=version,
            source=raw,
        )

    def serialize(self) -> bytes:
        if self.source is not None:
            return self.source
        out = bytearray(struct.pack("<I", self.version))
        for chunk in (
            self.package.serialize(),
            self.permissions,
            self.metadata.serialize(),
            self.bindings,
        ):
            out += struct.pack("<I", len(chunk)) + chunk
        return bytes(out)

    def touch(self) -> None:
        """Say that the blob no longer matches the bytes it was read from.

        The bindings go with it: they sign the package as it was, and
        Excel is content without them.
        """
        self.source = None
        self.bindings = b""
