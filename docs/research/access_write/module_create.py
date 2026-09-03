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
from vba_module_table import add_module, entries, entry_bytes, next_reserve  # noqa: E402
from vba_project_table import add_module_flag, append_identifier  # noqa: E402

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
#: One module's line in `Modules/PropData`, before its folder name.
FOLDER_ENTRY = bytes.fromhex("050902")
FOLDER_SUFFIX = "CB0".encode("utf-16-le")
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


def add_to_folder_list(payload: bytes, folder: str) -> bytes:
    """Append a folder to `Modules/PropData`."""
    return payload + FOLDER_ENTRY + folder.encode("utf-16-le") + FOLDER_SUFFIX


def next_folder(taken: set[str]) -> str:
    """The character after the highest folder in use, which is how Access
    numbered a third module `5` where the second was `4`."""
    highest = max((ord(name) for name in taken if len(name) == 1), default=ord("0") - 1)
    return chr(highest + 1)


def add_to_project(text: str, name: str) -> str:
    eol = chr(13) + chr(10)
    lines = text.split(eol)
    last = max(i for i, line in enumerate(lines) if line.startswith("Module="))
    lines.insert(last + 1, f"Module={name}")
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}=38, 38, 1786, 1030, ")
    return eol.join(lines)


# --- the module table in _VBA_PROJECT ---------------------------------------


def fresh_cookie(taken: set[bytes], template: bytes) -> bytes:
    """A module's 20-byte cookie is `<two characters><the project's own
    eight>`, and every module in a project has its own leading pair
    (measured: 08, 09 and 0@ against one suffix).  This keeps the
    project's half and takes the first pair nobody has."""
    suffix = template[4:]
    for code in range(0x30, 0x7F):
        candidate = ("0" + chr(code)).encode("utf-16-le") + suffix
        if candidate not in taken:
            return candidate
    raise LookupError("no unused module cookie left")


def add_to_vba_project(
    blob: bytes, cookie: bytes, stream_name: str, name: str, offset: int, module_cookie: bytes
) -> bytes:
    """Append an entry, giving it the reserve the project's trailer offers
    and a cookie no other module holds."""
    known = entries(blob, cookie)
    blob = add_module_flag(blob, len(known))
    blob, operand = append_identifier(blob, name)
    entry = entry_bytes(
        stream_name,
        fresh_cookie({e.cookie for e in known}, known[0].cookie),
        operand,
        name,
        module_cookie,
        next_reserve(blob, cookie),
        offset,
    )
    return add_module(blob, cookie, entry)


# --- the whole operation -----------------------------------------------------


def create(source: Path, target: Path, name: str, template: str = "Module1", seed: int = 0,
           skip: str = "", donor: Path | None = None) -> list[str]:
    skipped = set(skip.split(",")) if skip else set()
    db = AccessDatabase(source)
    done: list[str] = []
    rng = random.Random(seed)
    stream_name = stream_row_name(rng)

    # The compiled shape may come from another database, so a project can
    # be given a small empty module rather than a clone of whatever large
    # one it happens to hold.
    from_file = donor if donor is not None else source
    origin = next(
        (s for s in AccessReader(from_file).find_module_streams() if s.name.lower() == template.lower()), None
    )
    if origin is None:
        raise LookupError(f"no module named {template!r} in {from_file} to clone")

    storage = db.table(STORAGE)
    dir_rid, dir_stream = next(
        (rid, decompress(row["Lv"]))
        for rid, row in storage.rows_with_ids()
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes)
    )
    cookie = project_cookie(dir_stream)
    donor_dir = dir_stream if donor is None else _dir_stream_of(donor)
    # Every module carries its own MODULEEND2 word; two sharing one is
    # not something Access ever writes.
    module_word = rng.randbytes(2)
    template_stream = stream_name_of(donor_dir if donor is not None else dir_stream, template)
    template_offset = next(
        int.from_bytes(payload, "little")
        for _at, ident, _size, payload in records(dir_stream)
        if ident == MODULEOFFSET
    )

    perf = Perf(bytes(origin.raw), _offset_of(donor_dir, template))
    body = "\r\n".join(
        [f'Attribute VB_Name = "{name}"'] + perf.source_lines()
    ).encode("latin-1")
    row, offset = perf.build(new_source=body)

    ids = [r["Id"] for _rid, r in storage.rows_with_ids() if isinstance(r["Id"], int)]
    next_id = max(ids) + 1
    folders = {
        str(r["Name"]) for _rid, r in storage.rows_with_ids() if r["ParentId"] == MODULES_STORAGE and r["Type"] == 1
    }
    folder = next_folder(folders)
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
            storage.update_row(
                rid,
                {"Lv": add_to_vba_project(value, cookie, stream_name, name, offset, module_word)},
            )
            done.append("_VBA_PROJECT")
        elif r["Name"] == "\x03DirData" and r["ParentId"] == MODULES_STORAGE and "dirdata" not in skipped:
            storage.update_row(rid, {"Lv": add_to_dir_data(value, name)})
            done.append("DirData")
        elif r["Name"] == "PropData" and r["ParentId"] == MODULES_STORAGE:
            storage.update_row(rid, {"Lv": add_to_folder_list(value, folder)})
            done.append("Modules/PropData")
        elif r["Name"] == "PROJECTwm" and "wm" not in skipped:
            storage.update_row(rid, {"Lv": add_to_project_wm(value, name)})
            done.append("PROJECTwm")
        elif r["Name"] == "PROJECT" and "proj" not in skipped:
            storage.update_row(rid, {"Lv": add_to_project(value.decode("latin-1"), name).encode("latin-1")})
            done.append("PROJECT")

    if "dir" not in skipped:
        storage.update_row(
            dir_rid,
            {"Lv": compress(add_to_dir(dir_stream, dir_block(name, stream_name, offset, module_word)))},
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


def _dir_stream_of(path: Path) -> bytes:
    other = AccessDatabase(path)
    for _rid, row in other.table(STORAGE).rows_with_ids():
        if row["Name"] == "dir" and isinstance(row.get("Lv"), bytes):
            return decompress(row["Lv"])
    raise LookupError(f"{path} has no dir stream")


def _offset_of(stream: bytes, name: str) -> int:
    want, seen = name.encode("latin-1"), False
    for _at, ident, _size, payload in records(stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULEOFFSET and seen:
            return int.from_bytes(payload, "little")
    raise LookupError(f"no module offset for {name!r}")


if __name__ == "__main__":
    where = sys.argv[7] if len(sys.argv) > 7 else None
    print(
        create(
            Path(sys.argv[1]),
            Path(sys.argv[2]),
            sys.argv[3],
            *sys.argv[4:7],
            donor=Path(where) if where else None,
        )
    )
