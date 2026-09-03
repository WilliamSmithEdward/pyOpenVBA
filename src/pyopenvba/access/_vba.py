"""The VBA project inside a Jet 4 / ACE database.

Access does not keep its VBA in a CFB file the way the other Office hosts
do.  It keeps one ``MSysAccessStorage`` row per stream, under a small
tree of folders, and a module costs rows in five of them plus rows in
three catalog tables.  Every one of those places has to agree: a module
listed in one and missing from another is a module Access will show and
then refuse to open.

Writing takes the **source route**.  ``_VBA_PROJECT`` is [MS-OVBA]'s
PerformanceCache -- the compiled project -- and its ``Version`` field
says which build of VBA compiled it.  Write a version the host does not
recognise and VBA discards the cache and compiles the project from the
source in the module streams, which is what Access's own ``/decompile``
does.  So a module's stream here is the compressed source alone, with the
dir stream's MODULEOFFSET at zero, and none of the compiled tables have
to be generated.

One of them could not have been.  The 32-slot table ahead of the module
table in ``_VBA_PROJECT`` is runtime state rather than a function of the
file: adding the same module to the same database twice, in two Access
sessions, produces two different tables, with every per-module field in
the file identical between them.  ``docs/research/access_write`` keeps
the full record, and a byte-exact rename that leaves the cache intact.

The cost of the source route is that the next open recompiles.  The file
stops matching what Access wrote until Access rewrites it, and the
project's existing source has to compile -- a stale cache no longer hides
a module that does not.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass

from pyopenvba.access._storage import (
    PROP_DATA,
    STORAGE_TABLE,
    add_to_dir_data,
    dir_data_entries,
    next_folder,
    remove_from_dir_data,
    rename_dir_data,
    stream_row_name,
)
from pyopenvba.access_read import AccessError
from pyopenvba.vba import compress, decompress

__all__ = [
    "PROP_DATA",
    "STORAGE_TABLE",
    "add_to_dir_data",
    "dir_data_entries",
    "next_folder",
    "remove_from_dir_data",
    "rename_dir_data",
    "stream_row_name",
]


# --- the dir stream ----------------------------------------------------------
#: The one record whose size field is not a size.
PROJECTVERSION = 0x0009
PROJECTMODULES = 0x000F
TERMINATOR = 0x0010
MODULENAME = 0x0019
MODULESTREAMNAME = 0x001A
MODULEDOCSTRING = 0x001C
MODULEHELPCONTEXT = 0x001E
MODULETYPE_PROCEDURAL = 0x0021
MODULETYPE_CLASS = 0x0022
MODULEEND = 0x002B
MODULEEND2 = 0x002C
MODULEOFFSET = 0x0031
MODULESTREAMNAMEUNICODE = 0x0032
MODULENAMEUNICODE = 0x0047
MODULEDOCSTRINGUNICODE = 0x0048

#: What ``kind`` means in the dir stream.
MODULETYPE = {"module": MODULETYPE_PROCEDURAL, "class": MODULETYPE_CLASS}
KIND_OF_TYPE = {value: key for key, value in MODULETYPE.items()}

# --- the compiled cache ------------------------------------------------------
#: ``_VBA_PROJECT`` opens ``cc 61 <u16 Version> 00 <u16>``.
CACHE_SIGNATURE = bytes.fromhex("cc61")
CACHE_VERSION_AT = 2
#: Any value the host does not know will do; this one is one below the
#: version Access 2016 writes, so it can never collide with a real build.
STALE_VERSION = 0x0099

# --- source attributes -------------------------------------------------------
QUOTE = chr(34)
CRLF = chr(13) + chr(10)
#: Access's class-module base, measured off a class the VBE added.  A
#: class stream without it loads but will not instantiate.
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

# --- the storage rows a module occupies --------------------------------------
#: One folder's line in ``Modules/PropData``, before its name.
FOLDER_ENTRY = bytes.fromhex("050902")
FOLDER_SUFFIX = "CB0".encode("utf-16-le")

# --- the catalog -------------------------------------------------------------
OBJECT_MODULE = -32761
NAV_MODULE_TYPE = 32775
#: The navigation-pane group Access files modules under.
NAV_MODULE_GROUP = 8
#: Access hands out object ids four at a time: ``Module1`` at
#: -2147483640, then -2147483635, -2147483631, -2147483627, -2147483623 as
#: a project grew.  Taking max + 1 lands inside the range another object
#: holds, and ``AllModules(i).Name`` then fails.
OBJECT_ID_STEP = 4


@dataclass(frozen=True)
class VBAModule:
    """One module in the database's VBA project."""

    name: str
    kind: str
    stream_name: str
    source: str

    @property
    def is_class(self) -> bool:
        return self.kind == "class"


def records(stream: bytes) -> Iterator[tuple[int, int, int, bytes]]:
    """``(offset, id, size, payload)`` for each record of a decompressed
    dir stream."""
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


def module_blocks(dir_stream: bytes) -> list[tuple[str, str, str]]:
    """``(name, stream row name, kind)`` for every module the dir stream
    lists, in the order it lists them."""
    out: list[tuple[str, str, str]] = []
    name = stream_name = None
    for _at, ident, _size, payload in records(dir_stream):
        if ident == MODULENAME:
            name, stream_name = payload.decode("latin-1"), None
        elif ident == MODULESTREAMNAME:
            stream_name = payload.decode("latin-1")
        elif ident in KIND_OF_TYPE and name is not None and stream_name is not None:
            out.append((name, stream_name, KIND_OF_TYPE[ident]))
            name = stream_name = None
    return out


def stream_name_of(dir_stream: bytes, name: str) -> str:
    """The storage row a module's code lives in.

    Always through the dir stream: the row is named with 28 random
    capitals that have nothing to do with the module's name, and a module
    written source-only carries no p-code to scan the file for.
    """
    for module, stream_name, _kind in module_blocks(dir_stream):
        if module.lower() == name.lower():
            return stream_name
    raise AccessError(f"the VBA project has no module named {name!r}")


def module_offset_at(dir_stream: bytes, name: str) -> int:
    """Where a module's MODULEOFFSET payload starts."""
    want, seen = name.encode("latin-1"), False
    for at, ident, _size, payload in records(dir_stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULEOFFSET and seen:
            return at + 6
    raise AccessError(f"the dir stream has no MODULEOFFSET for {name!r}")


def _record(ident: int, payload: bytes) -> bytes:
    return ident.to_bytes(2, "little") + len(payload).to_bytes(4, "little") + payload


def dir_block(name: str, stream_name: str, cookie: bytes, kind: str) -> bytes:
    """The eleven records a module contributes to the dir stream."""
    return b"".join(
        (
            _record(MODULENAME, name.encode("latin-1")),
            _record(MODULENAMEUNICODE, name.encode("utf-16-le")),
            _record(MODULESTREAMNAME, stream_name.encode("latin-1")),
            _record(MODULESTREAMNAMEUNICODE, stream_name.encode("utf-16-le")),
            _record(MODULEDOCSTRING, b""),
            _record(MODULEDOCSTRINGUNICODE, b""),
            _record(MODULEOFFSET, bytes(4)),
            _record(MODULEHELPCONTEXT, bytes(4)),
            _record(MODULEEND2, cookie),
            _record(MODULETYPE[kind], b""),
            _record(MODULEEND, b""),
        )
    )


def _set_module_count(stream: bytes, delta: int) -> bytes:
    out = bytearray(stream)
    for at, ident, size, payload in records(bytes(out)):
        if ident == PROJECTMODULES and size == 2:
            count = int.from_bytes(payload, "little") + delta
            out[at + 6 : at + 8] = count.to_bytes(2, "little")
            break
    return bytes(out)


def add_to_dir(stream: bytes, block: bytes) -> bytes:
    """Insert a module's block before the terminator and count it."""
    at = None
    for offset, ident, _size, _payload in records(stream):
        if ident == TERMINATOR:
            at = offset
    if at is None:
        raise AccessError("the dir stream has no terminator")
    return _set_module_count(stream[:at] + block + stream[at:], 1)


def remove_from_dir(stream: bytes, name: str) -> bytes:
    """Drop a module's block and take one off the module count."""
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
    if start is None or end is None:
        raise AccessError(f"the dir stream has no module block for {name!r}")
    return _set_module_count(stream[:start] + stream[end:], -1)


def rename_in_dir(stream: bytes, old: str, new: str) -> bytes:
    """Rewrite a module's two name records."""
    out = bytearray(stream)
    for ident, encoding in ((MODULENAME, "latin-1"), (MODULENAMEUNICODE, "utf-16-le")):
        want, text = old.encode(encoding), new.encode(encoding)
        header = _record(ident, want)
        at = out.find(header)
        if at < 0:
            raise AccessError(f"the dir stream has no {ident:#06x} record for {old!r}")
        out[at : at + len(header)] = _record(ident, text)
    return bytes(out)


def set_module_offset(stream: bytes, name: str, offset: int) -> bytes:
    out = bytearray(stream)
    at = module_offset_at(bytes(out), name)
    out[at : at + 4] = offset.to_bytes(4, "little")
    return bytes(out)


# --- the compiled cache ------------------------------------------------------


def invalidate_cache(blob: bytes) -> bytes:
    """Mark the compiled project stale so VBA rebuilds it from source."""
    if blob[: len(CACHE_SIGNATURE)] != CACHE_SIGNATURE:
        raise AccessError("_VBA_PROJECT does not start with its signature")
    out = bytearray(blob)
    out[CACHE_VERSION_AT : CACHE_VERSION_AT + 2] = STALE_VERSION.to_bytes(2, "little")
    return bytes(out)


# --- a module's own stream ---------------------------------------------------


def attribute_lines(name: str, kind: str) -> list[str]:
    """The attributes a module's source opens with.  A class carries
    seven more, ``VB_Base`` among them."""
    lines = ["Attribute VB_Name = " + QUOTE + name + QUOTE]
    if kind == "class":
        lines += [f"Attribute {field} = {value}" for field, value in CLASS_ATTRIBUTES]
    return lines


def split_source(text: str) -> tuple[list[str], list[str]]:
    """A module's leading ``Attribute`` block and everything after it."""
    lines = text.split(CRLF)
    at = 0
    while at < len(lines) and lines[at].startswith("Attribute "):
        at += 1
    return lines[:at], lines[at:]


def module_stream(attributes: list[str], code: str) -> bytes:
    """A source-only module stream: the attributes, the body, compressed."""
    body = code.replace(CRLF, chr(10)).replace(chr(13), chr(10)).split(chr(10))
    return compress(CRLF.join(attributes + body).encode("latin-1"))


def read_source(stream: bytes, offset: int) -> str:
    """A module's source, from MODULEOFFSET on."""
    return decompress(stream[offset:]).decode("latin-1")


def rename_attribute(text: str, old: str, new: str) -> str:
    want = "Attribute VB_Name = " + QUOTE + old + QUOTE
    if want not in text:
        raise AccessError(f"the module holds no VB_Name attribute for {old!r}")
    return text.replace(want, "Attribute VB_Name = " + QUOTE + new + QUOTE)


# --- PROJECTwm and PROJECT --------------------------------------------------


def project_wm_entry(name: str) -> bytes:
    return name.encode("latin-1") + bytes(1) + name.encode("utf-16-le") + bytes(2)


def add_to_project_wm(payload: bytes, name: str) -> bytes:
    return payload[:-2] + project_wm_entry(name) + bytes(2)


def remove_from_project_wm(payload: bytes, name: str) -> bytes:
    want = project_wm_entry(name)
    if want not in payload:
        raise AccessError(f"PROJECTwm holds no entry for {name!r}")
    return payload.replace(want, b"")


def rename_project_wm(payload: bytes, old: str, new: str) -> bytes:
    want = project_wm_entry(old)
    if want not in payload:
        raise AccessError(f"PROJECTwm holds no entry for {old!r}")
    return payload.replace(want, project_wm_entry(new))


def add_to_project(text: str, name: str, kind: str) -> str:
    """Access lists a standard module as ``Module=`` and a class as
    ``Class=``, both in the same block, and gives each a window rectangle
    under ``[Workspace]``."""
    lines = text.split(CRLF)
    last = max(i for i, line in enumerate(lines) if line.startswith(("Module=", "Class=")))
    lines.insert(last + 1, ("Class=" if kind == "class" else "Module=") + name)
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}=38, 38, 1786, 1030, ")
    return CRLF.join(lines)


def remove_from_project(text: str, name: str) -> str:
    lines = [
        line
        for line in text.split(CRLF)
        if line not in (f"Module={name}", f"Class={name}")
        and not re.match(re.escape(name) + "=", line)
    ]
    return CRLF.join(lines)


def rename_project(text: str, old: str, new: str) -> str:
    """The ``Module=``/``Class=`` line and the ``[Workspace]`` line.  The
    stream's lines end CR LF, so the end anchor has to allow the CR."""
    tail = "(?=" + chr(92) + "r?$)"
    for keyword in ("Module", "Class"):
        text = re.sub(
            "(?m)^" + keyword + "=" + re.escape(old) + tail, keyword + "=" + new, text
        )
    return re.sub("(?m)^" + re.escape(old) + "=", new + "=", text)


# --- the folder list ---------------------------------------------------------


def add_to_folder_list(payload: bytes, folder: str) -> bytes:
    return payload + FOLDER_ENTRY + folder.encode("utf-16-le") + FOLDER_SUFFIX


def remove_from_folder_list(payload: bytes, folder: str) -> bytes:
    entry = FOLDER_ENTRY + folder.encode("utf-16-le") + FOLDER_SUFFIX
    return payload.replace(entry, b"", 1)
