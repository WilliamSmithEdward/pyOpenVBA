"""Walk a decompressed dir stream as [MS-OVBA] records."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.vba import decompress  # noqa: E402

# The one record whose size field is not a size.
PROJECTVERSION = 0x0009
NAMES = {
    0x0001: "PROJECTSYSKIND", 0x0002: "PROJECTLCID", 0x0003: "PROJECTCODEPAGE",
    0x0004: "PROJECTNAME", 0x0005: "PROJECTDOCSTRING", 0x0006: "PROJECTHELPFILEPATH",
    0x0007: "PROJECTHELPCONTEXT", 0x0008: "PROJECTLIBFLAGS", 0x0009: "PROJECTVERSION",
    0x000C: "PROJECTCONSTANTS", 0x000D: "REFERENCEREGISTERED", 0x000E: "REFERENCEPROJECT",
    0x000F: "PROJECTMODULES", 0x0010: "TERMINATOR", 0x0013: "PROJECTCOOKIE",
    0x0014: "REFERENCENAME", 0x0016: "REFERENCEORIGINAL", 0x0019: "MODULENAME",
    0x001A: "MODULESTREAMNAME", 0x001C: "MODULEDOCSTRING", 0x001E: "MODULEHELPCONTEXT",
    0x001F: "MODULECOOKIE", 0x0021: "MODULETYPE_PROCEDURAL", 0x0022: "MODULETYPE_DOCUMENT",
    0x0025: "MODULEREADONLY", 0x0028: "MODULEPRIVATE", 0x002B: "MODULEEND",
    0x002C: "MODULEEND2", 0x002F: "REFERENCECONTROL", 0x0030: "REFERENCEEXTENDED",
    0x0031: "MODULEOFFSET", 0x0032: "MODULESTREAMNAMEUNICODE", 0x0033: "REFERENCEORIGINAL2",
    0x003C: "PROJECTCOMPATVERSION", 0x0040: "MODULEDOCSTRINGUNICODE",
    0x0047: "MODULENAMEUNICODE", 0x0048: "PROJECTCONSTANTSUNICODE",
    0x004A: "PROJECTHELPFILEPATHUNICODE", 0x003D: "REFERENCENAMEUNICODE",
    0x003E: "REFERENCEORIGINALUNICODE",
}


def records(stream: bytes):
    """``(offset, id, size, payload)`` for each record."""
    at = 0
    while at + 6 <= len(stream):
        ident = int.from_bytes(stream[at : at + 2], "little")
        size = int.from_bytes(stream[at + 2 : at + 6], "little")
        if ident == PROJECTVERSION:
            yield at, ident, 6, stream[at + 2 : at + 12]
            at += 12
            continue
        if at + 6 + size > len(stream):
            return
        yield at, ident, size, stream[at + 6 : at + 6 + size]
        at += 6 + size


def dir_stream(path: Path) -> bytes:
    db = AccessDatabase(path)
    for _rid, row in db.table("MSysAccessStorage").rows_with_ids():
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes):
            return decompress(row["Lv"])
    raise LookupError("no dir stream")


if __name__ == "__main__":
    stream = dir_stream(Path(sys.argv[1]))
    print(f"{len(stream)} bytes")
    for at, ident, size, payload in records(stream):
        label = NAMES.get(ident, f"{ident:#06x}")
        text = payload[:40]
        print(f"  {at:6} {label:26} size {size:5}  {text!r}")
