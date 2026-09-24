"""A workbook Excel saved, made fit to commit as a test fixture.

Excel records the folder it saved a workbook in: an x15ac:absPath element,
wrapped in an mc:AlternateContent block, in xl/workbook.xml. That is a
folder on the machine that ran the measurement, so it comes out before a
workbook goes into tests/fixtures. Nothing else changes. Every other part
keeps its bytes, header and compressed body alike, through the library's
own package reader and writer, and xl/workbook.xml keeps its header and
compression method, deflated again as Office deflates.

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
#: The block Excel writes, and only it: an AlternateContent whose one choice is the save path.
SAVE_PATH = re.compile(r'<mc:AlternateContent\b[^>]*><mc:Choice Requires="x15"><x15ac:absPath\b[^>]*/>'
                       r"</mc:Choice></mc:AlternateContent>")
_STORED = 0


def without_save_path(package: bytes) -> bytes:
    """``package`` with the folder Excel saved it in taken out, or as it is where it records none."""
    opc = OpcFile.parse(package)
    if not opc.has(WORKBOOK_PART):
        return package
    entry = opc.entry(WORKBOOK_PART)
    xml = entry.read().decode("utf-8")
    kept, found = SAVE_PATH.subn("", xml)
    if found == 0:
        return package
    if found > 1:
        raise ValueError(f"{WORKBOOK_PART} records {found} save paths, where Excel writes one")
    data = kept.encode("utf-8")
    entry.body = data if entry.method == _STORED else raw_compress(data)
    entry.crc = zlib.crc32(data) & 0xFFFFFFFF
    entry.uncompressed_size = len(data)
    opc.source = None
    return opc.serialize()


def strip_save_path(path: Path) -> None:
    """Take the save path out of the workbook Excel saved at ``path``."""
    raw = path.read_bytes()
    clean = without_save_path(raw)
    if clean != raw:
        path.write_bytes(clean)


def copy_saved(source: Path, target: Path) -> None:
    """Copy the workbook Excel saved at ``source`` to ``target``, without its save path."""
    target.write_bytes(without_save_path(source.read_bytes()))
