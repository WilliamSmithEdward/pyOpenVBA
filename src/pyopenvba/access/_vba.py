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

import random
import re
from collections.abc import Iterator
from dataclasses import dataclass

from pyopenvba.access._props import PropertyValue
from pyopenvba.access._storage import (
    LIST_HEADER,
    PROP_DATA,
    STORAGE_TABLE,
    access_order,
    add_to_dir_data,
    dir_data_entries,
    next_folder,
    remove_from_dir_data,
    rename_dir_data,
    stream_row_name,
)
from pyopenvba.access_read import AccessError
from pyopenvba.vba import compress, decompress, encode_mbcs, encoding_for_codepage

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
PROJECTCODEPAGE = 0x0003
PROJECTNAME = 0x0004
PROJECTCOOKIE = 0x0013
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
#: One folder's line in ``Modules/PropData`` opens with this tag, then the
#: size of the rest of the line and the size of the folder's name.
FOLDER_TAG = 0x05
FOLDER_SUFFIX = "CB0".encode("utf-16-le")

# --- the catalog -------------------------------------------------------------
OBJECT_MODULE = -32761
NAV_MODULE_TYPE = 32775

# --- the project's own rows -------------------------------------------------
#: `VBA/AcessVBAData`, in Access's spelling: two words that are always 1,
#: then how many modules the project holds, the code behind a form or
#: report included.  Access keeps the count as modules come and go (13 in
#: a project of 13, 0 once the last is deleted).
VBA_DATA = "AcessVBAData"
EMPTY_VBA_DATA = bytes.fromhex("01000000" "01000000" "00000000")
#: The root `PropData` gains this property, the project's code page, when
#: a database gets its first VBA project.
PROJECT_CODE_PAGE_PROPERTY = 0x6A
ROOT_PROPERTY_TAG = 0x02
#: And MSysDb's properties gain these two, as Access 16 writes them:
#: `HasOfflineLists` 70 and `ProjVer` 142, both typed Integer, the second
#: four bytes long all the same.
PROJECT_DATABASE_PROPERTIES = {
    "HasOfflineLists": PropertyValue(3, 0, bytes.fromhex("4600")),
    "ProjVer": PropertyValue(3, 0, bytes.fromhex("8e000000")),
}
#: Where an Access 2000 file keeps its VBA project, a layout not read here.
ACCESS_2000_STORAGE = "MSysAccessObjects"


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


@dataclass(frozen=True)
class ModuleStream:
    """A module's stream as the project's storage holds it.

    ``data`` is the storage row's whole value: the compiled cache when the
    module has one, then from ``offset``, its MODULEOFFSET, the compressed
    source.  ``home`` is where that value starts, as ``(page, slot)``: its
    first long-value row, or the storage row itself when the value is
    short enough to sit inline.  A module whose row is missing has no data
    and no home.
    """

    name: str
    kind: str
    stream_name: str
    data: bytes
    offset: int
    encoding: str
    home: tuple[int, int] | None

    @property
    def text(self) -> str:
        """The source in the project's code page, ``Attribute`` lines
        included; empty when the row is missing."""
        return read_source(self.data, self.offset, self.encoding) if self.data else ""


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


def code_page(dir_stream: bytes) -> int:
    """The PROJECTCODEPAGE the dir stream declares.

    Falls back to 1252 for a stream that carries no such record, which is
    what the reader does ([MS-OVBA] 2.3.4.2.1.4 makes the record
    mandatory, so this is a damaged-file path rather than a real one).
    """
    for _at, ident, size, payload in records(dir_stream):
        if ident == PROJECTCODEPAGE and size >= 2:
            return int.from_bytes(payload[:2], "little")
    return 1252


def encoding_of(dir_stream: bytes) -> str:
    """The Python codec for whatever code page the project declares.

    Every ANSI string in the dir stream, the module streams and PROJECTwm
    is in this encoding, not latin-1.  The two agree on 0x00-0x7F and
    0xA0-0xFF and disagree on 0x80-0x9F, which is where cp1252 keeps the
    em dash, the curly quotes, the ellipsis and the euro sign, so reading
    a Western project as latin-1 is byte-lossless but silently wrong the
    moment the text is displayed or re-encoded (GitHub issue #18).
    """
    return encoding_for_codepage(code_page(dir_stream))


def module_blocks(dir_stream: bytes) -> list[tuple[str, str, str]]:
    """``(name, stream row name, kind)`` for every module the dir stream
    lists, in the order it lists them."""
    out: list[tuple[str, str, str]] = []
    encoding = encoding_of(dir_stream)
    name = stream_name = None
    for _at, ident, _size, payload in records(dir_stream):
        if ident == MODULENAME:
            name, stream_name = payload.decode(encoding, errors="replace"), None
        elif ident == MODULESTREAMNAME:
            stream_name = payload.decode(encoding, errors="replace")
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
    want, seen = encode_mbcs(name, encoding_of(dir_stream)), False
    for at, ident, _size, payload in records(dir_stream):
        if ident == MODULENAME:
            seen = payload == want
        elif ident == MODULEOFFSET and seen:
            return at + 6
    raise AccessError(f"the dir stream has no MODULEOFFSET for {name!r}")


def _record(ident: int, payload: bytes) -> bytes:
    return ident.to_bytes(2, "little") + len(payload).to_bytes(4, "little") + payload


def dir_block(name: str, stream_name: str, cookie: bytes, kind: str, encoding: str) -> bytes:
    """The eleven records a module contributes to the dir stream.

    ``encoding`` is the project's, from :func:`encoding_of`.  A character
    the code page cannot hold folds to ``?`` in the ANSI record and stays
    exact in the Unicode one beside it, which is what the VBE writes.
    """
    return b"".join(
        (
            _record(MODULENAME, encode_mbcs(name, encoding)),
            _record(MODULENAMEUNICODE, name.encode("utf-16-le")),
            _record(MODULESTREAMNAME, encode_mbcs(stream_name, encoding)),
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
    want = encode_mbcs(name, encoding_of(stream))
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
    ansi = encoding_of(stream)
    for ident, want, text in (
        (MODULENAME, encode_mbcs(old, ansi), encode_mbcs(new, ansi)),
        (MODULENAMEUNICODE, old.encode("utf-16-le"), new.encode("utf-16-le")),
    ):
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


def replace_record(stream: bytes, ident: int, payload: bytes) -> bytes:
    """The dir stream with its first ``ident`` record holding ``payload``."""
    for at, found, size, _old in records(stream):
        if found == ident:
            return stream[:at] + _record(ident, payload) + stream[at + 6 + size :]
    raise AccessError(f"the dir stream has no {ident:#06x} record")


# --- the project's own rows -------------------------------------------------


def with_module_count(payload: bytes, count: int) -> bytes:
    """``AcessVBAData`` counting ``count`` modules."""
    return payload[:8] + count.to_bytes(4, "little") + payload[12:]


def with_root_property(payload: bytes, ident: int, value: int) -> bytes:
    """The root ``PropData`` with property ``ident`` set, appended as
    ``02 <u32 id> <u32 value>`` when it is not there, which is how the
    code page arrives with a first project."""
    at = 4
    while at + 9 <= len(payload) and payload[at] == ROOT_PROPERTY_TAG:
        if int.from_bytes(payload[at + 1 : at + 5], "little") == ident:
            return payload[: at + 5] + value.to_bytes(4, "little") + payload[at + 9 :]
        at += 9
    entry = bytes((ROOT_PROPERTY_TAG,)) + ident.to_bytes(4, "little") + value.to_bytes(4, "little")
    return payload + entry


#: What a new project's ``CMG``, ``DPB`` and ``GC`` records hold, as every
#: project Access wrote here holds them: not protected, no password,
#: visible.
UNPROTECTED = {"CMG": bytes(4), "DPB": bytes(1), "GC": b"\xff"}


def encrypt_project_data(data: bytes, project_id: str, rng: random.Random) -> str:
    """One of ``PROJECT``'s ``CMG``, ``DPB`` and ``GC`` values ([MS-OVBA]
    2.4.3.2): a random seed, then each byte XORed with a running sum that
    starts from the project key, the byte sum of the ``ID`` text with its
    braces.  The key is what ties the record to its project -- decrypted
    against the ID of every project Access wrote here, the records carried
    exactly that sum -- and VBA takes a project whose key does not match
    for a protected one."""
    seed = rng.randrange(256)
    key = sum(project_id.encode("ascii")) & 0xFF
    version_enc, key_enc = seed ^ 2, seed ^ key
    out = bytearray((seed, version_enc, key_enc))
    plain_prev, enc1, enc2 = key, key_enc, version_enc
    ignored = rng.randbytes((seed & 6) // 2)
    for byte in ignored + len(data).to_bytes(4, "little") + data:
        byte_enc = byte ^ ((enc2 + plain_prev) & 0xFF)
        out.append(byte_enc)
        enc2, enc1, plain_prev = enc1, byte_enc, byte
    return out.hex().upper()


def new_project_text(text: str, name: str, project_id: str, rng: random.Random) -> str:
    """A ``PROJECT`` stream for a new project, from one Access wrote with
    its modules taken out: the opening block's ``ID`` and ``Name`` lines
    rewritten, and its ``CMG``, ``DPB`` and ``GC`` encrypted afresh against
    the new ID, the rest as it stood."""
    lines = text.split(CRLF)
    for i, line in enumerate(lines):
        if line.startswith("["):
            break
        key = line.split("=", 1)[0]
        if key == "ID":
            lines[i] = "ID=" + QUOTE + project_id + QUOTE
        elif key == "Name":
            lines[i] = "Name=" + QUOTE + name + QUOTE
        elif key in UNPROTECTED:
            lines[i] = key + "=" + QUOTE + encrypt_project_data(UNPROTECTED[key], project_id, rng) + QUOTE
    return CRLF.join(lines)


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


def module_stream(attributes: list[str], code: str, encoding: str) -> bytes:
    """A source-only module stream: the attributes, the body, compressed.

    ``encoding`` is the project's, from :func:`encoding_of`.  Source is
    stored in the code page, so an em dash in a Western project is one
    byte here and not the three UTF-8 would take.
    """
    body = code.replace(CRLF, chr(10)).replace(chr(13), chr(10)).split(chr(10))
    return compress(encode_mbcs(CRLF.join(attributes + body), encoding))


def read_source(stream: bytes, offset: int, encoding: str) -> str:
    """A module's source, from MODULEOFFSET on."""
    return decompress(stream[offset:]).decode(encoding, errors="replace")


def rename_attribute(text: str, old: str, new: str) -> str:
    want = "Attribute VB_Name = " + QUOTE + old + QUOTE
    if want not in text:
        raise AccessError(f"the module holds no VB_Name attribute for {old!r}")
    return text.replace(want, "Attribute VB_Name = " + QUOTE + new + QUOTE)


# --- PROJECTwm and PROJECT --------------------------------------------------
#: How PROJECT names the module behind a form or report: the keyword, the
#: module's name, a slash, and a flag word Access owns.  A design's module
#: is listed this way and never as ``Module=`` or ``Class=``.
DOC_CLASS = "DocClass"
#: The three keywords that open the module block.
MODULE_KEYWORDS = ("Module=", "Class=", DOC_CLASS + "=")


def project_wm_entry(name: str, encoding: str) -> bytes:
    return encode_mbcs(name, encoding) + bytes(1) + name.encode("utf-16-le") + bytes(2)


def add_to_project_wm(payload: bytes, name: str, encoding: str) -> bytes:
    return payload[:-2] + project_wm_entry(name, encoding) + bytes(2)


def remove_from_project_wm(payload: bytes, name: str, encoding: str) -> bytes:
    want = project_wm_entry(name, encoding)
    if want not in payload:
        raise AccessError(f"PROJECTwm holds no entry for {name!r}")
    return payload.replace(want, b"")


def rename_project_wm(payload: bytes, old: str, new: str, encoding: str) -> bytes:
    want = project_wm_entry(old, encoding)
    if want not in payload:
        raise AccessError(f"PROJECTwm holds no entry for {old!r}")
    return payload.replace(want, project_wm_entry(new, encoding))


def add_to_project(text: str, name: str, kind: str) -> str:
    """Access lists a standard module as ``Module=`` and a class as
    ``Class=``, both in the same block, and gives each a window rectangle
    under ``[Workspace]``.

    The block can be empty -- delete a project's last module and there is
    no line to sit after -- and Access opens it right below the ``ID=``
    line, which is where the first one goes.
    """
    lines = text.split(CRLF)
    _list_module(lines, ("Class=" if kind == "class" else "Module=") + name)
    _add_workspace_line(lines, f"{name}=38, 38, 1786, 1030, ")
    return CRLF.join(lines)


def _list_module(lines: list[str], entry: str) -> None:
    """After the block's last module line, or right below ``ID=`` when the
    block is empty."""
    listed = [i for i, line in enumerate(lines) if line.startswith(MODULE_KEYWORDS)]
    if listed:
        lines.insert(max(listed) + 1, entry)
    else:
        opener = next((i for i, line in enumerate(lines) if line.startswith("ID=")), -1)
        lines.insert(opener + 1, entry)


def _add_workspace_line(lines: list[str], line: str) -> None:
    """A window rectangle under ``[Workspace]``.  A project that has never
    held a module has no such section, and Access opens one for the first
    module or form code, after a blank line."""
    if not any(existing.strip() == "[Workspace]" for existing in lines):
        lines[-1:-1] = ["", "[Workspace]"]
    lines.insert(len(lines) - 1, line)


def without_empty_workspace(text: str) -> str:
    """``PROJECT`` without an empty ``[Workspace]`` section, as Access
    writes a project it made for a form, report or macro rather than a
    module."""
    lines = text.split(CRLF)
    if lines[-3:] == ["", "[Workspace]", ""]:
        lines[-3:] = [""]
    return CRLF.join(lines)


def remove_from_project(text: str, name: str) -> str:
    """Every line naming a module: its ``Module=``, ``Class=`` or
    ``DocClass=`` line, and its ``[Workspace]`` rectangle."""
    lines = [
        line
        for line in text.split(CRLF)
        if line not in (f"Module={name}", f"Class={name}")
        and not line.startswith(f"{DOC_CLASS}={name}/")
        and not re.match(re.escape(name) + "=", line)
    ]
    return CRLF.join(lines)


def rename_project(text: str, old: str, new: str) -> str:
    """The ``Module=``, ``Class=`` or ``DocClass=`` line and the
    ``[Workspace]`` line.  The stream's lines end CR LF, so the end anchor
    has to allow the CR.

    A design's module is listed only as ``DocClass=<name>/<flags>``, never
    as ``Module=`` or ``Class=``.  Renaming without reaching it leaves a
    DocClass naming a module the project no longer has, which Access
    reports as a corrupt project on the first VBE reference (GitHub issue
    #21).  The flag word after the slash is Access's, so the match stops
    at it and it is carried through.
    """
    quoted = re.escape(old)
    tail = "(?=" + chr(92) + "r?$)"
    for keyword in ("Module", "Class"):
        text = re.sub("(?m)^" + keyword + "=" + quoted + tail, keyword + "=" + new, text)
    text = re.sub("(?m)^" + DOC_CLASS + "=" + quoted + "(?=/)", DOC_CLASS + "=" + new, text)
    return re.sub("(?m)^" + quoted + "=", new + "=", text)


# --- the folder list ---------------------------------------------------------


def folder_entry(folder: str) -> bytes:
    """A folder's line in the list: ``05 09 02 "4" "CB0"`` for folder `4`,
    and ``05 0b 04 "10" "CB0"`` for folder `10`, both as Access wrote them.
    The two sizes are what grow with the name."""
    name = folder.encode("utf-16-le")
    return bytes((FOLDER_TAG, 1 + len(name) + len(FOLDER_SUFFIX), len(name))) + name + FOLDER_SUFFIX


def _folder_list_parts(payload: bytes) -> tuple[list[str], dict[str, bytes], bytes]:
    """The folders in stored order, each folder's whole line, and whatever
    follows the last line."""
    folders: list[str] = []
    lines: dict[str, bytes] = {}
    at = LIST_HEADER
    while at + 3 <= len(payload) and payload[at] == FOLDER_TAG:
        end = at + 2 + payload[at + 1]
        folder = payload[at + 3 : at + 3 + payload[at + 2]].decode("utf-16-le")
        folders.append(folder)
        lines[folder] = payload[at:end]
        at = end
    return folders, lines, payload[at:]


def _in_folder_list(payload: bytes, folder: str, *, again: bool) -> bytes:
    """The list with `folder`'s line erased, and inserted again when
    `again`, in the order Access writes it after (see `access_order`).  A
    list without the line is left alone, as Access leaves it."""
    folders, lines, tail = _folder_list_parts(payload)
    if folder not in lines:
        return payload
    added = (folder,) if again else ()
    order = access_order(folders, remove=(folder,), add=added)
    return payload[:LIST_HEADER] + b"".join(lines[name] for name in order) + tail


def remove_from_folder_list(payload: bytes, folder: str) -> bytes:
    return _in_folder_list(payload, folder, again=False)


def refile_in_folder_list(payload: bytes, folder: str) -> bytes:
    """A renamed object's line: Access erases it and inserts it again, so
    it moves though the folder keeps its name."""
    return _in_folder_list(payload, folder, again=True)


# --- code behind a form or report ---------------------------------------------
# A document module belongs to its design, not to `Modules`: it has no
# storage folder, no `MSysObjects` row and no entry in the container's own
# lists.  What it does have is a stream of its own, a dir block, a
# `PROJECTwm` entry, and a `DocClass=` line in `PROJECT` -- and without
# that last one Access loads the module but the form does not answer to
# it, which is the whole difference between a class module and this.
DOC_CLASS_SUFFIX = "/&H00000000"
#: The window rectangle Access gives a document module.
DOC_WORKSPACE = "0, 0, 0, 0, C"
#: A document module's attributes: creatable and predeclared, where a
#: plain class module is neither, and a `VB_Base` naming a CLSID the
#: design's own `TypeInfo` repeats.
DOCUMENT_ATTRIBUTES = (
    ("VB_GlobalNameSpace", "False"),
    ("VB_Creatable", "True"),
    ("VB_PredeclaredId", "True"),
    ("VB_Exposed", "False"),
    ("VB_TemplateDerived", "False"),
    ("VB_Customizable", "False"),
)
#: Where the design's `TypeInfo` keeps that CLSID, and where its folder's
#: `PropData` records that it has a module at all.
TYPE_INFO_CLSID = 16
PROP_DATA_HAS_MODULE = 9


def document_attributes(name: str, clsid: str) -> list[str]:
    """The attributes the module behind a form or report opens with."""
    return [
        "Attribute VB_Name = " + QUOTE + name + QUOTE,
        "Attribute VB_Base = " + QUOTE + "0{" + clsid + "}" + QUOTE,
        *(f"Attribute {field} = {value}" for field, value in DOCUMENT_ATTRIBUTES),
    ]


def add_to_project_documents(text: str, name: str) -> str:
    """A `DocClass=` line, and a window rectangle under `[Workspace]`."""
    lines = text.split(CRLF)
    _list_module(lines, f"{DOC_CLASS}={name}{DOC_CLASS_SUFFIX}")
    _add_workspace_line(lines, f"{name}={DOC_WORKSPACE}")
    return CRLF.join(lines)


def remove_from_project_documents(text: str, name: str) -> str:
    """A document module's ``DocClass=`` line and its workspace rectangle.

    Matched on the prefix, because the flag word after the slash and the
    rectangle are Access's and a database it has edited need not carry the
    ones written here.
    """
    return CRLF.join(
        line
        for line in text.split(CRLF)
        if not line.startswith(f"{DOC_CLASS}={name}/") and not line.startswith(f"{name}=")
    )
