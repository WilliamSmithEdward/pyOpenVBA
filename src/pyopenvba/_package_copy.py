"""Write a zip-based Office file with a few parts changed and the rest as they were stored.

Saving a VBA project changes one part, ``vbaProject.bin``, and at most a
handful of small parts beside it. Reading and recompressing everything
else costs time proportional to the expanded content. Copying unchanged
compressed bytes avoids that work, using the same container writer as
Power Query and the document surfaces. It still costs time and memory
proportional to the compressed package size.
A changed part is deflated by zlib at its default level, as ``zipfile``
deflates it, so it is stored as the same bytes as before this module
existed.

:func:`copy_package` answers None for a package it would not carry over
faithfully, and the caller then writes it with ``zipfile`` as it always
has.  That covers ZIP64, a part compressed some other way than deflate
that has to be written anew, an encrypted part, a part name that is not
UTF-8, an unknown compression method, unsupported patched-data or
strong-encryption flags, and a package naming one part twice. It also
covers a package whose two directories disagree -- a part's own header naming it
differently, giving it other sizes in a ZIP64 record, or placing it over
another part or the central directory -- which ``zipfile`` refuses,
and so still does.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from collections.abc import Mapping, Sequence
from dataclasses import replace

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._opc import Entry, OpcFile

_STORED = 0
_DEFLATED = 8
#: General purpose bit 0: the part is encrypted.
_ENCRYPTED = 0x0001
#: Patched data and strong encryption are refused by zipfile when reading.
_UNSUPPORTED_FLAGS = 0x0060
_UTF8 = 0x0800
#: General purpose bit 3: the sizes follow the part's bytes instead of standing in its header.
_DESCRIPTOR = 0x0008
#: The date ``zipfile`` gives a part with none, 1980-01-01, as [APPNOTE] 4.4.6 packs one.
_NO_DATE = 0x0021
#: A size or offset of this much is held in a ZIP64 record instead.
_ZIP64 = 0xFFFFFFFF
#: The extra field record that holds one.
_ZIP64_RECORD = 0x0001


def _has_zip64(extra: bytes) -> bool:
    """Whether an extra field holds a ZIP64 record, or cannot be read to say."""
    at = 0
    while at < len(extra):
        if at + 4 > len(extra):
            return True
        kind, size = struct.unpack_from("<HH", extra, at)
        if kind == _ZIP64_RECORD:
            return True
        at += 4 + size
    return False


def _pack(data: bytes, method: int) -> bytes | None:
    """``data`` as a part stored by ``method`` holds it; None for a method not written here."""
    if method == _STORED:
        return data
    if method == _DEFLATED:
        deflater = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION, zlib.DEFLATED, -15)
        return deflater.compress(data) + deflater.flush()
    return None


def copy_package(
    raw: bytes,
    infos: Sequence[zipfile.ZipInfo],
    edits: Mapping[str, bytes | None],
    added: Mapping[str, bytes],
) -> bytes | None:
    """The package ``raw`` with ``edits`` applied and ``added`` appended, every other part as it was stored.

    ``infos`` are the package's entries as ``zipfile`` read them, which
    is what the edits are named by.  ``edits`` gives a part its new
    bytes, or None to leave it out; a part is kept in its place, under
    its own compression method, dates and attributes.  ``added`` are
    parts the package does not have yet, deflated and put last.

    None means the package is one this does not write; see the module's
    docstring.
    """
    try:
        package = OpcFile.parse(raw)
        central_start = struct.unpack_from("<I", raw, raw.rfind(b"PK\x05\x06") + 16)[0]
    except (PowerQueryError, UnicodeDecodeError, struct.error):
        return None
    if len(package.entries) != len(infos) or len({info.filename for info in infos}) != len(infos):
        return None

    entries: list[Entry] = []
    # Where the last part's bytes ended: each part has to start past the one before it.
    end = 0
    for entry, info in zip(package.entries, infos, strict=True):
        # The two readers have to agree on what the part is, or an edit could land on the wrong one.
        if (
            entry.name != info.filename
            or len(entry.body) != info.compress_size
            or entry.uncompressed_size != info.file_size
            or entry.crc != info.CRC
            or entry.method not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED, zipfile.ZIP_BZIP2, zipfile.ZIP_LZMA)
            or entry.flags & (_ENCRYPTED | _UNSUPPORTED_FLAGS)
            or info.compress_size >= _ZIP64
            or info.file_size >= _ZIP64
            or _has_zip64(entry.local_extra)
            or _has_zip64(entry.central_extra)
        ):
            return None
        # So do the part's own header and the directory: on its name, and on parts not lying over each other.
        name = entry.name.encode("utf-8")
        named = info.header_offset + 30
        local_name_length, local_extra_length = struct.unpack_from("<HH", raw, info.header_offset + 26)
        if (
            local_name_length != len(name)
            or local_extra_length != len(entry.local_extra)
            or info.header_offset < end
            or raw[named : named + local_name_length + local_extra_length] != name + entry.local_extra
        ):
            return None
        local_flags = struct.unpack_from("<H", raw, info.header_offset + 6)[0]
        if (local_flags ^ entry.flags) & _UTF8:
            return None
        end = named + local_name_length + local_extra_length + len(entry.body)
        if end > central_start:
            return None
        # The sizes are written in the header, so nothing follows the part's bytes to say them again.
        flags = entry.flags & ~_DESCRIPTOR
        if entry.name in edits:
            data = edits[entry.name]
            if data is None:
                continue
            body = _pack(data, entry.method)
            if body is None:
                return None
            entry = replace(
                entry, body=body, flags=flags, crc=zlib.crc32(data) & 0xFFFFFFFF, uncompressed_size=len(data)
            )
        elif flags != entry.flags:
            entry = replace(entry, flags=flags)
        entries.append(entry)

    for name, data in added.items():
        body = _pack(data, _DEFLATED)
        assert body is not None
        entries.append(
            Entry(
                name=name,
                body=body,
                method=_DEFLATED,
                flags=0,
                dos_time=0,
                dos_date=_NO_DATE,
                crc=zlib.crc32(data) & 0xFFFFFFFF,
                uncompressed_size=len(data),
            )
        )

    try:
        return OpcFile(entries=entries).serialize()
    except struct.error:
        # Past what a ZIP holds without ZIP64: more than 65535 parts, or 4 GB.
        return None
