"""Create a module through the storage engine.

Measured against Access's own VBComponents.Add on a one-module project.
What a new module costs the file:

* three new MSysAccessStorage rows -- a numbered storage folder under
  `Modules`, a 13-byte `PropData` under it, and the module's own stream
  under `VBA`, whose row name is 28 random capitals
* a dir block of eleven records, MODULENAME through MODULEEND, and
  PROJECTMODULES up by one
* an entry each in DirData, PROJECTwm and PROJECT
* an MSysObjects row of type -32761 under the Modules container, and a
  matching MSysNavPaneObjectIDs row
* the version word of `_VBA_PROJECT` set to something VBA does not
  recognise, which is what makes the rest of it unnecessary

`_VBA_PROJECT` is a performance cache and its `Version` field says which
build of VBA wrote it.  Change that word and VBA throws the cache away
and compiles the project from the source in the module streams, which is
the same thing Access's own `/decompile` does.  So a new module needs no
p-code at all: its stream is the compressed source with MODULEOFFSET 0,
and the compiled tables -- the module table, the identifier table, the
per-module records and the two hash tables -- are rebuilt by VBA.

Modelling those tables instead was the long way round, and one of them
cannot be modelled: the 32-slot table before the module table is runtime
state, not a function of the file.  Adding the same module to the same
database twice, in two Access sessions, gives two different tables.
"""

from __future__ import annotations

import datetime as dt
import random
import string
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
for _where in (ROOT / "src", ROOT / "docs/research/access_write", ROOT / "docs/research/pcode"):
    if str(_where) not in sys.path:
        sys.path.insert(0, str(_where))

from dir_records import records  # noqa: E402
from pyopenvba.access import AccessDatabase  # noqa: E402
from pyopenvba.vba import compress, decompress  # noqa: E402
from module_rename import _dir_data_entry, _project_wm_entry, drop_srp  # noqa: E402

STORAGE = "MSysAccessStorage"
QUOTE = chr(34)
CRLF = chr(13) + chr(10)
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
#: The dir stream's MODULETYPE, and the only thing in the file that says
#: whether a module is a class besides its attributes and its PROJECT line.
MODULETYPE = {"module": 0x0021, "class": 0x0022}
#: Access's own class-module base, measured off a class the VBE added.
CLASS_BASE = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
CLASS_ATTRIBUTES = (
    ("VB_Base", QUOTE + CLASS_BASE + QUOTE),
    ("VB_GlobalNameSpace", "False"),
    ("VB_Creatable", "False"),
    ("VB_PredeclaredId", "False"),
    ("VB_Exposed", "False"),
    ("VB_TemplateDerived", "False"),
    ("VB_Customizable", "False"),
)
PROP_DATA = bytes.fromhex("00000000020000000000000000")
#: One module's line in `Modules/PropData`, before its folder name.
FOLDER_ENTRY = bytes.fromhex("050902")
FOLDER_SUFFIX = "CB0".encode("utf-16-le")
COOKIE_BYTES = 20


def stream_row_name(seed: random.Random, taken: set[str] = frozenset()) -> str:
    """28 random capitals, and no row under `VBA` may already hold it."""
    while True:
        name = "".join(seed.choice(string.ascii_uppercase) for _ in range(28))
        if name not in taken:
            return name


# --- the dir stream ----------------------------------------------------------


def _record(ident: int, payload: bytes) -> bytes:
    return ident.to_bytes(2, "little") + len(payload).to_bytes(4, "little") + payload


def dir_block(name: str, stream_name: str, offset: int, cookie: bytes, kind: str = "module") -> bytes:
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
            _record(MODULETYPE[kind], b""),
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


def add_to_project(text: str, name: str, kind: str = "module") -> str:
    """Access lists a standard module as `Module=` and a class as
    `Class=`, both in the same block."""
    eol = CRLF
    lines = text.split(eol)
    last = max(i for i, line in enumerate(lines) if line.startswith(("Module=", "Class=")))
    lines.insert(last + 1, ("Class=" if kind == "class" else "Module=") + name)
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}=38, 38, 1786, 1030, ")
    return eol.join(lines)


# --- the performance cache ---------------------------------------------------
#: `_VBA_PROJECT` opens with `cc 61 <u16 version> 00 <u16>`, and the
#: version is the build of VBA that compiled the cache.  Any value the
#: host does not know makes it recompile from source.
VERSION_AT = 2
STALE_VERSION = 0x0099


def invalidate_cache(blob: bytes) -> bytes:
    """Mark the compiled project stale so VBA rebuilds it on load."""
    if blob[:2] != bytes.fromhex("cc61"):
        raise ValueError("_VBA_PROJECT does not start with its signature")
    out = bytearray(blob)
    out[VERSION_AT : VERSION_AT + 2] = STALE_VERSION.to_bytes(2, "little")
    return bytes(out)


def attribute_lines(name: str, kind: str = "module") -> list[str]:
    """The attributes a module's source opens with.  A class carries
    seven more, `VB_Base` among them -- without it the VBE reads the
    stream but Access will not instantiate the class."""
    lines = ["Attribute VB_Name = " + QUOTE + name + QUOTE]
    if kind == "class":
        lines += [f"Attribute {field} = {value}" for field, value in CLASS_ATTRIBUTES]
    return lines


def module_source(name: str, code: str, kind: str = "module") -> bytes:
    """A module stream is the compressed source, and the source opens
    with the attributes that carry the module's name and its kind."""
    body = code.replace(CRLF, chr(10)).replace(chr(13), chr(10))
    return compress(CRLF.join(attribute_lines(name, kind) + body.split(chr(10))).encode("latin-1"))


# --- the whole operation -----------------------------------------------------


def create(source: Path, target: Path, name: str, code: str = "Option Compare Database",
           kind: str = "module", seed: int | None = None) -> list[str]:
    """Add a module holding `code` and return what was written.  `kind` is
    `"module"` or `"class"`."""
    if kind not in MODULETYPE:
        raise ValueError(f"kind must be one of {sorted(MODULETYPE)}, not {kind!r}")
    db = AccessDatabase(source)
    done: list[str] = []
    rng = random.Random(seed)
    row = module_source(name, code, kind)

    storage = db.table(STORAGE)
    stream_name = stream_row_name(
        rng, {str(r["Name"]) for _rid, r in storage.rows_with_ids() if r["ParentId"] == VBA_STORAGE}
    )
    dir_rid, dir_stream = next(
        (rid, decompress(row_["Lv"]))
        for rid, row_ in storage.rows_with_ids()
        if row_["Name"] == "dir" and isinstance(row_.get("Lv"), bytes)
    )
    # Every module carries its own MODULEEND2 word; two sharing one is
    # not something Access ever writes.
    module_word = rng.randbytes(2)

    folders = {
        str(r["Name"]) for _rid, r in storage.rows_with_ids() if r["ParentId"] == MODULES_STORAGE and r["Type"] == 1
    }
    folder = next_folder(folders)
    stamp = dt.datetime.now().replace(microsecond=0)

    # Ids come from the table's own AutoNumber rather than from max + 1:
    # every database Access wrote has the counter equal to its highest id,
    # and leaving it behind makes Access's next insert collide.
    rid = storage.insert_row(
        {"ParentId": MODULES_STORAGE, "Name": folder, "Type": 1, "DateCreate": stamp, "DateUpdate": stamp}
    )
    folder_id = next(r["Id"] for at, r in storage.rows_with_ids() if at == rid)
    storage.insert_row(
        {"ParentId": folder_id, "Name": "PropData", "Type": 2, "Lv": PROP_DATA,
         "DateCreate": stamp, "DateUpdate": stamp}
    )
    storage.insert_row(
        {"ParentId": VBA_STORAGE, "Name": stream_name, "Type": 2, "Lv": row,
         "DateCreate": stamp, "DateUpdate": stamp}
    )
    done.append(f"storage folder {folder!r}, PropData and stream {stream_name!r}")

    for rid, r in list(storage.rows_with_ids()):
        value = r.get("Lv")
        if not isinstance(value, bytes) or not value:
            continue
        if r["Name"] == "_VBA_PROJECT":
            storage.update_row(rid, {"Lv": invalidate_cache(value)})
            done.append("_VBA_PROJECT marked stale")
        elif r["Name"] == chr(3) + "DirData" and r["ParentId"] == MODULES_STORAGE:
            storage.update_row(rid, {"Lv": add_to_dir_data(value, name)})
            done.append("DirData")
        elif r["Name"] == "PropData" and r["ParentId"] == MODULES_STORAGE:
            storage.update_row(rid, {"Lv": add_to_folder_list(value, folder)})
            done.append("Modules/PropData")
        elif r["Name"] == "PROJECTwm":
            storage.update_row(rid, {"Lv": add_to_project_wm(value, name)})
            done.append("PROJECTwm")
        elif r["Name"] == "PROJECT":
            storage.update_row(
                rid, {"Lv": add_to_project(value.decode("latin-1"), name, kind).encode("latin-1")}
            )
            done.append("PROJECT")

    storage.update_row(
        dir_rid,
        {"Lv": compress(add_to_dir(dir_stream, dir_block(name, stream_name, 0, module_word, kind)))},
    )
    done.append("dir")

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


if __name__ == "__main__":
    if len(sys.argv) < 4:
        raise SystemExit("usage: module_create.py SOURCE TARGET NAME [CODE] [module|class]")
    body = sys.argv[4] if len(sys.argv) > 4 else "Option Compare Database"
    print(
        create(
            Path(sys.argv[1]),
            Path(sys.argv[2]),
            sys.argv[3],
            body.replace("|", chr(10)),
            sys.argv[5] if len(sys.argv) > 5 else "module",
        )
    )
