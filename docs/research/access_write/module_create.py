"""Create a module through the storage engine.

Measured against Access's own VBComponents.Add on a one-module project.
What a new module costs the file:

* three new MSysAccessStorage rows -- a numbered storage folder under
  `Modules`, a 13-byte `PropData` under it, and the module's own stream
  under `VBA`, whose row name is 28 random capitals
* a dir block of eleven records, MODULENAME through MODULEEND, and
  PROJECTMODULES up by one
* an entry each in DirData, PROJECTwm and PROJECT
* a `_VBA_PROJECT` module entry, its count up by one, and an identifier
  record for the name
* an MSysObjects row of type -32761 under the Modules container, and a
  matching MSysNavPaneObjectIDs row

The compiled shape is cloned from a module that already exists and has no
procedures of its own -- the template's `Module1` is one -- because
synthesising a procedure table from nothing is a separate problem.
"""

from __future__ import annotations

import datetime as dt
import random
import string
import sys
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from accdb_write import Perf  # noqa: E402
from module_delete import project_cookie, stream_name_of  # noqa: E402
from dir_records import records  # noqa: E402
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.access_read import AccessReader  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402
from module_rename import _dir_data_entry, _project_wm_entry, drop_srp  # noqa: E402
from vba_module_table import module_count  # noqa: E402
from vba_project_table import append_identifier  # noqa: E402

STORAGE = "MSysAccessStorage"
MODULE_TYPE = -32761
NAV_MODULE_TYPE = 32775
MODULES_STORAGE = 6
VBA_STORAGE = 17
MODULENAME = 0x0019
MODULEEND = 0x002B
MODULEEND2 = 0x002C
MODULEOFFSET = 0x0031
PROJECTMODULES = 0x000F
TERMINATOR = 0x0010
PROP_DATA = bytes.fromhex("00000000020000000000000000")
COOKIE_BYTES = 20


def stream_row_name(seed: random.Random) -> str:
    return "".join(seed.choice(string.ascii_uppercase) for _ in range(28))


# --- the dir stream ----------------------------------------------------------


def _record(ident: int, payload: bytes) -> bytes:
    return ident.to_bytes(2, "little") + len(payload).to_bytes(4, "little") + payload


def dir_block(name: str, stream_name: str, offset: int, cookie: bytes) -> bytes:
    return b"".join(
        (
            _record(MODULENAME, name.encode("latin-1")),
            _record(0x0047, name.encode("utf-16-le")),
            _record(0x001A, stream_name.encode("latin-1")),
            _record(0x0032, stream_name.encode("utf-16-le")),
            _record(0x001C, b""),
            _record(0x0048, b""),
            _record(MODULEOFFSET, offset.to_bytes(4, "little")),
            _record(0x001E, bytes(4)),
            _record(MODULEEND2, cookie),
            _record(0x0021, b""),
            _record(MODULEEND, b""),
        )
    )


def add_to_dir(stream: bytes, block: bytes) -> bytes:
    at = None
    for offset, ident, _size, _payload in records(stream):
        if ident == TERMINATOR:
            at = offset
    if at is None:
        raise LookupError("the dir stream has no terminator")
    out = bytearray(stream[:at] + block + stream[at:])
    for offset, ident, size, payload in records(bytes(out)):
        if ident == PROJECTMODULES and size == 2:
            out[offset + 6 : offset + 8] = (int.from_bytes(payload, "little") + 1).to_bytes(2, "little")
            break
    return bytes(out)


def module_cookie(stream: bytes, name: str) -> bytes:
    want, seen = name.encode("latin-1"), False
    for _at, ident, _size, payload in records(stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULEEND2 and seen:
            return payload
    raise LookupError(f"no module cookie for {name!r}")


# --- the small lists ---------------------------------------------------------


def add_to_dir_data(payload: bytes, name: str) -> bytes:
    tail = bytes((4, 0, 0, 0))
    body = payload[: -len(tail)] + bytes(4) if payload.endswith(tail) else payload
    return body + _dir_data_entry(name) + tail


def add_to_project_wm(payload: bytes, name: str) -> bytes:
    return payload[:-2] + _project_wm_entry(name) + bytes(2)


def add_to_project(text: str, name: str) -> str:
    eol = chr(13) + chr(10)
    lines = text.split(eol)
    last = max(i for i, line in enumerate(lines) if line.startswith("Module="))
    lines.insert(last + 1, f"Module={name}")
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}=38, 38, 1786, 1030, ")
    return eol.join(lines)


# --- the module table in _VBA_PROJECT ---------------------------------------


@dataclass
class Entry:
    start: int
    end: int
    cookie: bytes  # the 20-byte per-module string
    tail: bytes  # ff ff and the 17 bytes that close an entry


def read_entry(blob: bytes, stream_name: str, name: str) -> Entry:
    stream_bytes = stream_name.encode("utf-16-le")
    at = blob.find(stream_bytes)
    if at < 0:
        raise LookupError(f"_VBA_PROJECT has no entry for stream {stream_name!r}")
    cookie_at = at + len(stream_bytes) + 2
    text = name.encode("utf-16-le")
    where = blob.find(text, cookie_at)
    if where < 0:
        raise LookupError(f"_VBA_PROJECT has no name record for {name!r}")
    end = where + len(text) + 19
    return Entry(at - 2, end, blob[cookie_at : cookie_at + COOKIE_BYTES], blob[where + len(text) : end])


def add_to_vba_project(
    blob: bytes, cookie: bytes, template: Entry, stream_name: str, name: str, offset: int, module_cookie: bytes
) -> bytes:
    """Append an entry built from the template's fixed parts."""
    blob, operand = append_identifier(blob, name)
    stream_text, text = stream_name.encode("utf-16-le"), name.encode("utf-16-le")
    tail = bytearray(template.tail)
    tail[2:4] = module_cookie
    # The template's own module carries 0x0208 here; every module added
    # after it carries 0x0278 (measured on a three-module project, where
    # the second and third shared it).
    tail[10:12] = (0x0278).to_bytes(2, "little")
    tail[-4:] = offset.to_bytes(4, "little")
    entry = (
        len(stream_text).to_bytes(2, "little")
        + stream_text
        + COOKIE_BYTES.to_bytes(2, "little")
        + template.cookie
        + b"\xff\xff"
        + operand.to_bytes(2, "little")
        + len(text).to_bytes(2, "little")
        + text
        + bytes(tail)
    )
    at = template.end
    out = bytearray(blob[:at] + b"\xff\xff" + entry + blob[at:])
    where = module_count(bytes(out), cookie)
    out[where : where + 2] = (int.from_bytes(out[where : where + 2], "little") + 1).to_bytes(2, "little")
    return bytes(out)


# --- the whole operation -----------------------------------------------------


def create(source: Path, target: Path, name: str, template: str = "Module1", seed: int = 0, skip: str = "") -> list[str]:
    skipped = set(skip.split(",")) if skip else set()
    db = AccessDatabase(source)
    done: list[str] = []
    rng = random.Random(seed)
    stream_name = stream_row_name(rng)

    origin = next(
        (s for s in AccessReader(source).find_module_streams() if s.name.lower() == template.lower()), None
    )
    if origin is None:
        raise LookupError(f"no module named {template!r} to clone")

    storage = db.table(STORAGE)
    dir_rid, dir_stream = next(
        (rid, decompress(row["Lv"]))
        for rid, row in storage.rows_with_ids()
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes)
    )
    cookie = project_cookie(dir_stream)
    # Every module carries its own MODULEEND2 word; two sharing one is
    # not something Access ever writes.
    fresh_cookie = rng.randbytes(2)
    template_stream = stream_name_of(dir_stream, template)
    template_offset = next(
        int.from_bytes(payload, "little")
        for _at, ident, _size, payload in records(dir_stream)
        if ident == MODULEOFFSET
    )

    perf = Perf(bytes(origin.raw), _offset_of(dir_stream, template))
    body = "\r\n".join(
        [f'Attribute VB_Name = "{name}"'] + perf.source_lines()
    ).encode("latin-1")
    row, offset = perf.build(new_source=body)

    ids = [r["Id"] for _rid, r in storage.rows_with_ids() if isinstance(r["Id"], int)]
    next_id = max(ids) + 1
    folders = {
        str(r["Name"]) for _rid, r in storage.rows_with_ids() if r["ParentId"] == MODULES_STORAGE and r["Type"] == 1
    }
    folder = str(next(n for n in range(100) if str(n) not in folders))
    stamp = dt.datetime.now().replace(microsecond=0)

    if "rows" in skipped:
        db.save(target)
        return done
    if "folder" not in skipped:
        storage.insert_row(
            {"Id": next_id, "ParentId": MODULES_STORAGE, "Name": folder, "Type": 1, "DateCreate": stamp, "DateUpdate": stamp}
        )
        storage.insert_row(
            {"Id": next_id + 1, "ParentId": next_id, "Name": "PropData", "Type": 2, "Lv": PROP_DATA,
             "DateCreate": stamp, "DateUpdate": stamp}
        )
    if "stream" not in skipped:
        storage.insert_row(
            {"Id": next_id + 2, "ParentId": VBA_STORAGE, "Name": stream_name, "Type": 2, "Lv": row,
             "DateCreate": stamp, "DateUpdate": stamp}
        )
    done.append(f"storage folder {folder!r}, PropData and stream {stream_name!r}")

    entry = None
    for rid, r in list(storage.rows_with_ids()):
        value = r.get("Lv")
        if not isinstance(value, bytes) or not value:
            continue
        if r["Name"] == "_VBA_PROJECT" and "vba" not in skipped:
            entry = read_entry(value, template_stream, template)
            storage.update_row(
                rid,
                {"Lv": add_to_vba_project(value, cookie, entry, stream_name, name, offset,
                                          fresh_cookie)},
            )
            done.append("_VBA_PROJECT")
        elif r["Name"] == "\x03DirData" and _dir_data_entry(template) in value and "dirdata" not in skipped:
            storage.update_row(rid, {"Lv": add_to_dir_data(value, name)})
            done.append("DirData")
        elif r["Name"] == "PROJECTwm" and "wm" not in skipped:
            storage.update_row(rid, {"Lv": add_to_project_wm(value, name)})
            done.append("PROJECTwm")
        elif r["Name"] == "PROJECT" and "proj" not in skipped:
            storage.update_row(rid, {"Lv": add_to_project(value.decode("latin-1"), name).encode("latin-1")})
            done.append("PROJECT")

    if "dir" not in skipped:
        storage.update_row(
            dir_rid,
            {"Lv": compress(add_to_dir(dir_stream, dir_block(name, stream_name, offset, fresh_cookie)))},
        )
        done.append("dir")

    if "catalog" in skipped:
        db.save(target)
        return done
    objects = db.table("MSysObjects")
    container = next(e.id for e in db.catalog() if e.name == "Modules" and e.type == 3)
    owner = next(e.owner for e in db.catalog() if e.type == MODULE_TYPE and e.owner)
    object_id = max((e.id for e in db.catalog() if e.id < 0), default=-(2**31)) + 1
    objects.insert_row(
        {"Id": object_id, "ParentId": container, "Name": name, "Type": MODULE_TYPE, "Flags": 0,
         "Owner": owner, "DateCreate": stamp, "DateUpdate": stamp}
    )
    db.table("MSysNavPaneObjectIDs").insert_row({"Id": object_id, "Name": name, "Type": NAV_MODULE_TYPE})
    done.append(f"MSysObjects and the navigation pane (id {object_id})")

    done.append(f"dropped {drop_srp(db)} __SRP_ rows")
    db.save(target)
    return done


def _offset_of(stream: bytes, name: str) -> int:
    want, seen = name.encode("latin-1"), False
    for _at, ident, _size, payload in records(stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULEOFFSET and seen:
            return int.from_bytes(payload, "little")
    raise LookupError(f"no module offset for {name!r}")


if __name__ == "__main__":
    print(create(Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3], *sys.argv[4:]))
