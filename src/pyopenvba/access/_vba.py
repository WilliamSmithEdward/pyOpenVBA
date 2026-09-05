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
    entry = ("Class=" if kind == "class" else "Module=") + name
    listed = [i for i, line in enumerate(lines) if line.startswith(MODULE_KEYWORDS)]
    if listed:
        lines.insert(max(listed) + 1, entry)
    else:
        opener = next((i for i, line in enumerate(lines) if line.startswith("ID=")), -1)
        lines.insert(opener + 1, entry)
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}=38, 38, 1786, 1030, ")
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


def add_to_folder_list(payload: bytes, folder: str) -> bytes:
    return payload + FOLDER_ENTRY + folder.encode("utf-16-le") + FOLDER_SUFFIX


def remove_from_folder_list(payload: bytes, folder: str) -> bytes:
    entry = FOLDER_ENTRY + folder.encode("utf-16-le") + FOLDER_SUFFIX
    return payload.replace(entry, b"", 1)


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
    last = max(i for i, line in enumerate(lines) if line.startswith(MODULE_KEYWORDS))
    lines.insert(last + 1, f"{DOC_CLASS}={name}{DOC_CLASS_SUFFIX}")
    if any(line.strip() == "[Workspace]" for line in lines):
        lines.insert(len(lines) - 1, f"{name}={DOC_WORKSPACE}")
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


# --- the project's references -------------------------------------------------
# Three records each, in the dir stream ahead of PROJECTMODULES:
#
#     REFERENCEORIGINAL          the library's name, MBCS
#     REFERENCEORIGINALUNICODE   the same in UTF-16
#     REFERENCEREGISTERED        <u32 length> <libid> 00*6
#
# and the libid reads
#
#     *\G{GUID}#<major>.<minor>#<lcid>#<path>#<description>
#
# with the version in **hex**: DAO 12.0 is written `#c.0#`.  Access keeps
# references nowhere else -- `PROJECT` carries no `Reference=` line -- and
# the two every project has, VBA and Access itself, are not in the dir
# stream at all.
REFERENCEREGISTERED = 0x000D
REFERENCEORIGINAL = 0x0016
REFERENCEORIGINALUNICODE = 0x003E
LIBID_PREFIX = "*" + chr(92) + "G"
REFERENCE_TRAILER = bytes(6)


@dataclass(frozen=True)
class Reference:
    """One library the project points at."""

    name: str
    libid: str

    def _parts(self) -> list[str]:
        return self.libid.split("#")

    @property
    def guid(self) -> str:
        head = self._parts()[0]
        return head[len(LIBID_PREFIX) :] if head.startswith(LIBID_PREFIX) else head

    @property
    def version(self) -> tuple[int, int]:
        """The major and minor the libid names, which it writes in hex."""
        parts = self._parts()
        if len(parts) < 2 or "." not in parts[1]:
            return (0, 0)
        major, minor = parts[1].split(".", 1)
        try:
            return int(major, 16), int(minor, 16)
        except ValueError:
            return (0, 0)

    @property
    def path(self) -> str:
        parts = self._parts()
        return parts[3] if len(parts) > 3 else ""

    @property
    def description(self) -> str:
        parts = self._parts()
        return parts[4] if len(parts) > 4 else ""


def make_libid(guid: str, major: int, minor: int, path: str, description: str, lcid: int = 0) -> str:
    """The string a REFERENCEREGISTERED record holds."""
    if not guid.startswith("{"):
        guid = "{" + guid + "}"
    return f"{LIBID_PREFIX}{guid}#{major:x}.{minor:x}#{lcid}#{path}#{description}"


def references(dir_stream: bytes) -> list[Reference]:
    """Every library the dir stream points at, in its order."""
    out: list[Reference] = []
    encoding = encoding_of(dir_stream)
    name = ""
    for _at, ident, _size, payload in records(dir_stream):
        if ident == REFERENCEORIGINAL:
            name = payload.decode(encoding, errors="replace")
        elif ident == REFERENCEREGISTERED:
            length = int.from_bytes(payload[:4], "little")
            out.append(Reference(name, payload[4 : 4 + length].decode(encoding, errors="replace")))
    return out


def reference_block(name: str, libid: str, encoding: str) -> bytes:
    """The three records one reference contributes."""
    text = encode_mbcs(libid, encoding)
    return b"".join(
        (
            _record(REFERENCEORIGINAL, encode_mbcs(name, encoding)),
            _record(REFERENCEORIGINALUNICODE, name.encode("utf-16-le")),
            _record(
                REFERENCEREGISTERED,
                len(text).to_bytes(4, "little") + text + REFERENCE_TRAILER,
            ),
        )
    )


def add_reference(dir_stream: bytes, name: str, libid: str) -> bytes:
    """Insert a reference ahead of the modules, where Access keeps them."""
    at = next(
        (offset for offset, ident, _size, _payload in records(dir_stream) if ident == PROJECTMODULES),
        None,
    )
    if at is None:
        raise AccessError("the dir stream has no PROJECTMODULES record")
    block = reference_block(name, libid, encoding_of(dir_stream))
    return dir_stream[:at] + block + dir_stream[at:]


def remove_reference(dir_stream: bytes, name: str) -> bytes:
    """Drop a reference's three records."""
    want, start, end = encode_mbcs(name, encoding_of(dir_stream)), None, None
    for at, ident, size, payload in records(dir_stream):
        if ident == REFERENCEORIGINAL and payload == want:
            start = at
        elif start is not None and ident == REFERENCEREGISTERED and end is None:
            end = at + 6 + size
    if start is None or end is None:
        raise AccessError(f"the project has no reference named {name!r}")
    return dir_stream[:start] + dir_stream[end:]
