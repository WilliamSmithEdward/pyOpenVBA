"""Delete a module through the storage engine.

Measured against Access's own DoCmd.DeleteObject on a two-module project:

* the dir stream loses the module's whole record block, MODULENAME through
  MODULEEND, and PROJECTMODULES drops by one
* DirData, PROJECTwm and PROJECT lose their entries for it
* `_VBA_PROJECT` loses the module's UTF-16 entry and its module count
  drops by one, while the identifier table keeps the name and its two
  counters do not move
* its MSysAccessStorage stream row, its MSysObjects row and its
  MSysNavPaneObjectIDs row all go
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from dir_records import records  # noqa: E402
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access_read import AccessReader  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402
from vba_module_table import remove_module  # noqa: E402
from module_rename import _dir_data_entry, _project_wm_entry, drop_srp  # noqa: E402

STORAGE = "MSysAccessStorage"
MODULE_TYPE = -32761
MODULENAME = 0x0019
MODULEEND = 0x002B
PROJECTMODULES = 0x000F
MODULESTREAMNAME = 0x001A
PROJECTCOOKIE = 0x0013


def stream_name_of(stream: bytes, name: str) -> str:
    """The storage row a module's code lives in, from its dir block."""
    want, seen = name.encode("latin-1"), False
    for _at, ident, _size, payload in records(stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULESTREAMNAME and seen:
            return payload.decode("latin-1")
    raise LookupError(f"the dir stream names no stream for {name!r}")


def project_cookie(stream: bytes) -> bytes:
    for _at, ident, _size, payload in records(stream):
        if ident == PROJECTCOOKIE:
            return payload
    raise LookupError("the dir stream has no project cookie")


def remove_from_dir(stream: bytes, name: str) -> bytes:
    """Drop the module's record block and take one off PROJECTMODULES."""
    want = name.encode("latin-1")
    start = end = None
    for at, ident, _size, payload in records(stream):
        if ident == MODULENAME:
            if payload == want:
                start = at
            elif start is not None and end is None:
                end = at
        elif ident == MODULEEND and start is not None and end is None and at > start:
            end = at + 6
    if start is None:
        raise LookupError(f"the dir stream has no module block for {name!r}")
    out = bytearray(stream[:start] + stream[end:])
    for at, ident, size, payload in records(bytes(out)):
        if ident == PROJECTMODULES and size == 2:
            count = int.from_bytes(payload, "little")
            out[at + 6 : at + 8] = (count - 1).to_bytes(2, "little")
            break
    return bytes(out)


def remove_from_project(text: str, name: str) -> str:
    lines = [
        line
        for line in text.split(chr(13) + chr(10))
        if line != f"Module={name}" and not re.match(rf"^{re.escape(name)}=", line)
    ]
    return (chr(13) + chr(10)).join(lines)


def delete(source: Path, target: Path, name: str) -> list[str]:
    db = AccessDatabase(source)
    done: list[str] = []

    stream = next(
        (s for s in AccessReader(source).find_module_streams() if s.name.lower() == name.lower()), None
    )
    if stream is None:
        raise LookupError(f"no module stream named {name!r}")
    payload = bytes(stream.raw)

    storage = db.table(STORAGE)
    dir_stream = next(
        decompress(row["Lv"])
        for _rid, row in storage.rows_with_ids()
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes)
    )
    stream_name = stream_name_of(dir_stream, name)
    cookie = project_cookie(dir_stream)

    for rid, row in list(storage.rows_with_ids()):
        value = row.get("Lv")
        if not isinstance(value, bytes) or not value:
            continue
        if value == payload:
            storage.delete_row(rid, retire_empty=False)
            done.append("module stream row")
        elif row["Name"] == "dir":
            storage.update_row(rid, {"Lv": compress(remove_from_dir(decompress(value), name))})
            done.append("dir")
        elif row["Name"] == "\x03DirData" and _dir_data_entry(name) in value:
            # Each entry is followed by four bytes that go with it:
            # Access left `... Module1 00000000` where the two-name row
            # had read `... Module1 00000000 04 0e Alpha 04000000`.
            entry = _dir_data_entry(name)
            at = value.find(entry)
            storage.update_row(rid, {"Lv": value[:at] + value[at + len(entry) + 4 :]})
            done.append("DirData")
        elif row["Name"] == "PROJECTwm" and _project_wm_entry(name) in value:
            storage.update_row(rid, {"Lv": value.replace(_project_wm_entry(name), b"")})
            done.append("PROJECTwm")
        elif row["Name"] == "_VBA_PROJECT":
            storage.update_row(rid, {"Lv": remove_module(value, cookie, stream_name, name)})
            done.append("_VBA_PROJECT")
        elif row["Name"] == "PROJECT":
            text = value.decode("latin-1")
            fixed = remove_from_project(text, name)
            if fixed != text:
                storage.update_row(rid, {"Lv": fixed.encode("latin-1")})
                done.append("PROJECT")

    objects = db.table("MSysObjects")
    for rid, row in list(objects.rows_with_ids()):
        if row["Type"] == MODULE_TYPE and row["Name"] == name:
            objects.delete_row(rid, retire_empty=False)
            done.append("MSysObjects")
    nav = db.table("MSysNavPaneObjectIDs")
    for rid, row in list(nav.rows_with_ids()):
        if row["Name"] == name:
            nav.delete_row(rid, retire_empty=False)
            done.append("MSysNavPaneObjectIDs")

    done.append(f"dropped {drop_srp(db)} __SRP_ rows")
    db.save(target)
    return done


if __name__ == "__main__":
    print(delete(Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3]))
