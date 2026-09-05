"""The .NET binary serialization primitives the mashup metadata uses.

``QueryGroups`` is stored as a base64 string whose bytes come from
.NET's ``BinaryWriter``: a string carries a 7-bit-encoded length ahead of
its UTF-8 bytes, and a ``Guid`` is written in the mixed-endian layout
.NET uses.  Both were read off Microsoft's own serializer
(``QueriesMetadataSerializer.SerializeQueryGroups``) rather than guessed;
``tests/test_powerquery_groups.py`` holds the samples it produced.
"""

from __future__ import annotations

import struct
import uuid

from pyopenvba.exceptions import PowerQueryError

#: A 7-bit-encoded length never runs past five bytes for a 32-bit count.
_MAX_LENGTH_BYTES = 5


class BinaryReader:
    """Reads what .NET's ``BinaryWriter`` wrote."""

    def __init__(self, data: bytes) -> None:
        self._data = data
        self._at = 0

    @property
    def at(self) -> int:
        return self._at

    @property
    def remaining(self) -> int:
        return len(self._data) - self._at

    def _take(self, count: int) -> bytes:
        if count < 0 or self._at + count > len(self._data):
            raise PowerQueryError(
                f"this value ends after {self.remaining} bytes, and {count} were needed"
            )
        chunk = self._data[self._at : self._at + count]
        self._at += count
        return chunk

    def uint32(self) -> int:
        return int(struct.unpack("<I", self._take(4))[0])

    def int32(self) -> int:
        return int(struct.unpack("<i", self._take(4))[0])

    def byte(self) -> int:
        return self._take(1)[0]

    def boolean(self) -> bool:
        value = self.byte()
        if value > 1:
            raise PowerQueryError(f"a boolean byte is 0 or 1, not {value}")
        return value == 1

    def length(self) -> int:
        """The 7-bit-encoded length that precedes a string."""
        value = 0
        for step in range(_MAX_LENGTH_BYTES):
            piece = self.byte()
            value |= (piece & 0x7F) << (7 * step)
            if not piece & 0x80:
                return value
        raise PowerQueryError("a string length ran past five bytes")

    def text(self) -> str:
        return self._take(self.length()).decode("utf-8")

    def guid(self) -> uuid.UUID:
        return uuid.UUID(bytes_le=self._take(16))


class BinaryWriter:
    """Writes what .NET's ``BinaryReader`` expects."""

    def __init__(self) -> None:
        self._out = bytearray()

    def uint32(self, value: int) -> None:
        self._out += struct.pack("<I", value)

    def int32(self, value: int) -> None:
        self._out += struct.pack("<i", value)

    def byte(self, value: int) -> None:
        self._out.append(value)

    def boolean(self, value: bool) -> None:  # noqa: FBT001 - mirrors the wire shape
        self._out.append(1 if value else 0)

    def length(self, value: int) -> None:
        if value < 0:
            raise PowerQueryError("a string length cannot be negative")
        while value >= 0x80:
            self._out.append((value & 0x7F) | 0x80)
            value >>= 7
        self._out.append(value)

    def text(self, value: str) -> None:
        raw = value.encode("utf-8")
        self.length(len(raw))
        self._out += raw

    def guid(self, value: uuid.UUID) -> None:
        self._out += value.bytes_le

    def bytes(self) -> bytes:
        return bytes(self._out)
