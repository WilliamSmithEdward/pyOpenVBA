"""A workbook Excel saved, made fit to commit as a test fixture or bake as a template.

Excel records the folder it saved a workbook in: an x15ac:absPath element,
wrapped in an mc:AlternateContent block, in xl/workbook.xml; and in a .xlsb
a BrtAbsPath15 record, wrapped in the BrtACBegin and BrtACEnd records that
are the binary AlternateContent, in xl/workbook.bin. That is a folder on
the machine that saved it, so the block comes out before a workbook is
committed. Nothing else changes. Every other part keeps its bytes, header
and compressed body alike, through the library's own package reader and
writer; the workbook part keeps its header and compression method,
deflated again as Office deflates, and in a .xlsb every other record keeps
its bytes.

A measurement script copies a workbook Excel saved elsewhere into the
fixtures with copy_saved, and cleans one Excel saved there with
strip_save_path.
"""

from __future__ import annotations

import re
import zlib
from pathlib import Path

from pyopenvba._deflate import raw_compress
from pyopenvba.powerquery._opc import OpcFile

WORKBOOK_PART = "xl/workbook.xml"
BINARY_WORKBOOK_PART = "xl/workbook.bin"
#: The block Excel writes, and only it: an AlternateContent whose one choice is the save path.
SAVE_PATH = re.compile(r'<mc:AlternateContent\b[^>]*><mc:Choice Requires="x15"><x15ac:absPath\b[^>]*/>'
                       r"</mc:Choice></mc:AlternateContent>")
#: [MS-XLSB] record types: BrtACBegin, BrtACEnd and the BrtAbsPath15 between them.
_AC_BEGIN, _AC_END, _ABS_PATH = 0x0025, 0x0026, 0x0817
_STORED = 0


def records(data: bytes) -> list[tuple[int, int, int]]:
    """Each BIFF12 record of ``data`` as (start, end, type), read by [MS-XLSB] 2.1.4.

    A record opens with its type in one or two bytes and its size in one
    to four, seven bits a byte, the high bit saying another byte follows.
    """
    found: list[tuple[int, int, int]] = []
    at = 0
    while at < len(data):
        start = at
        kind = data[at] & 0x7F
        if data[at] & 0x80:
            at += 1
            kind |= (data[at] & 0x7F) << 7
        at += 1
        size = 0
        for shift in (0, 7, 14, 21):
            byte = data[at]
            at += 1
            size |= (byte & 0x7F) << shift
            if not byte & 0x80:
                break
        at += size
        found.append((start, at, kind))
    if at != len(data):
        raise ValueError(f"the records run {at - len(data)} bytes past the part's end")
    return found


def _without_xml_save_path(data: bytes) -> bytes | None:
    kept, found = SAVE_PATH.subn("", data.decode("utf-8"))
    if found > 1:
        raise ValueError(f"{WORKBOOK_PART} records {found} save paths, where Excel writes one")
    return kept.encode("utf-8") if found else None


def _without_binary_save_path(data: bytes) -> bytes | None:
    listed = records(data)
    at = [index for index, (_, _, kind) in enumerate(listed) if kind == _ABS_PATH]
    if not at:
        return None
    if len(at) > 1:
        raise ValueError(f"{BINARY_WORKBOOK_PART} records {len(at)} save paths, where Excel writes one")
    index = at[0]
    if not 0 < index < len(listed) - 1 or (listed[index - 1][2], listed[index + 1][2]) != (_AC_BEGIN, _AC_END):
        raise ValueError(f"{BINARY_WORKBOOK_PART} holds its save path outside the block Excel writes")
    return data[:listed[index - 1][0]] + data[listed[index + 1][1]:]


def without_save_path(package: bytes) -> bytes:
    """``package`` with the folder Excel saved it in taken out, or as it is where it records none.

    A file that is no zip, a binary .xls say, has no workbook part and
    comes back as it is.
    """
    if not package.startswith(b"PK\x03\x04"):
        return package
    opc = OpcFile.parse(package)
    for part, strip in ((WORKBOOK_PART, _without_xml_save_path), (BINARY_WORKBOOK_PART, _without_binary_save_path)):
        if not opc.has(part):
            continue
        entry = opc.entry(part)
        data = strip(entry.read())
        if data is None:
            return package
        entry.body = data if entry.method == _STORED else raw_compress(data)
        entry.crc = zlib.crc32(data) & 0xFFFFFFFF
        entry.uncompressed_size = len(data)
        opc.source = None
        return opc.serialize()
    return package


def strip_save_path(path: Path) -> None:
    """Take the save path out of the workbook Excel saved at ``path``."""
    raw = path.read_bytes()
    clean = without_save_path(raw)
    if clean != raw:
        path.write_bytes(clean)


def copy_saved(source: Path, target: Path) -> None:
    """Copy the workbook Excel saved at ``source`` to ``target``, without its save path."""
    target.write_bytes(without_save_path(source.read_bytes()))
