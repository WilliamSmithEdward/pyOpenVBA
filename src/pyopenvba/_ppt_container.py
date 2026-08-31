"""
Binary PowerPoint (``.ppt``) VBA container: locate, extract and write back
the embedded VBA project storage.

Unlike ``.doc`` and ``.xls``, where the whole file is a CFB whose root
holds the VBA storage, a ``.ppt`` keeps its VBA project *inside* the
``PowerPoint Document`` stream.  That stream is a flat sequence of records
addressed through a persist model:

- the ``Current User`` stream's ``CurrentUserAtom`` points at the newest
  ``UserEditAtom``,
- each ``UserEditAtom`` points at a ``PersistDirectoryAtom`` and at the
  previous edit,
- the directories map persist object ids to absolute byte offsets in the
  ``PowerPoint Document`` stream.

The VBA project is an ``ExOleObjStg`` record (0x1011) holding a whole CFB,
usually zlib-deflated and truncated without a proper stream end.  Its
persist id is named by the ``DocumentContainer``'s ``VbaInfoAtom``.

Writing back replaces that record in place.  Every structure carrying an
absolute offset -- the persist directories, the user-edit chain and the
``CurrentUserAtom`` -- is shifted by the size delta; nothing else in the
file addresses by offset, which is what the persist model is for.

Private module: the public surface is :class:`pyopenvba.PowerPointFile`.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass

from pyopenvba.cfb import CFB
from pyopenvba.exceptions import VBAProjectError

CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

DOC_STREAM = "PowerPoint Document"
CURRENT_USER_STREAM = "Current User"

RT_DOCUMENT_CONTAINER = 0x03E8
RT_VBA_INFO = 0x03FF
RT_VBA_INFO_ATOM = 0x0400
RT_USER_EDIT_ATOM = 0x0FF5
RT_PERSIST_DIRECTORY_ATOM = 0x1772
RT_EX_OLE_OBJ_STG = 0x1011
RT_CURRENT_USER_ATOM = 0x0FF6

# A record header is version/instance (2), type (2), payload length (4).
HEADER_SIZE = 8
# recVer 0xF marks a container record, whose payload is more records.
RECVER_CONTAINER = 0xF
# CurrentUserAtom payload: size(4) headerToken(4) offsetToCurrentEdit(4).
OFFSET_TO_CURRENT_EDIT = HEADER_SIZE + 8
# UserEditAtom payload: lastSlideIdRef(4) version(4) offsetLastEdit(4)
# offsetPersistDirectory(4) ...
OFFSET_LAST_EDIT = 8
OFFSET_PERSIST_DIRECTORY = 12
# A PersistDirectoryAtom entry packs the first persist id in the low 20
# bits and the run length in the top 12.
PERSIST_ID_MASK = 0xFFFFF
PERSIST_COUNT_SHIFT = 20


@dataclass(frozen=True)
class Record:
    """One record header inside the ``PowerPoint Document`` stream."""

    ver_instance: int
    rec_type: int
    start: int
    """Offset of the 8-byte header itself."""
    end: int
    """Offset just past the record's payload."""

    @property
    def rec_instance(self) -> int:
        return self.ver_instance >> 4

    @property
    def is_container(self) -> bool:
        return (self.ver_instance & 0xF) == RECVER_CONTAINER


def read_header(stream: bytes, offset: int) -> Record | None:
    """Parse a record header, or ``None`` if it does not fit the stream."""
    if offset < 0 or offset + HEADER_SIZE > len(stream):
        return None
    ver_instance, rec_type, rec_len = struct.unpack_from("<HHI", stream, offset)
    if rec_len > len(stream) - (offset + HEADER_SIZE):
        return None
    return Record(ver_instance, rec_type, offset, offset + HEADER_SIZE + rec_len)


def top_level_records(stream: bytes) -> list[Record]:
    """Every record at the top level of the stream, in file order."""
    records: list[Record] = []
    offset = 0
    while True:
        record = read_header(stream, offset)
        if record is None:
            return records
        records.append(record)
        offset = record.end


def find_descendant(
    doc: bytes, container: Record, rec_type: int
) -> Record | None:
    """Depth-first search for the first child record of ``rec_type``."""
    cursor = container.start + HEADER_SIZE
    while cursor + HEADER_SIZE <= container.end:
        child = read_header(doc, cursor)
        if child is None or child.end > container.end:
            return None
        if child.rec_type == rec_type:
            return child
        if child.is_container:
            nested = find_descendant(doc, child, rec_type)
            if nested is not None:
                return nested
        cursor = child.end
    return None


# ---------------------------------------------------------------------------
# Persist machinery
# ---------------------------------------------------------------------------

def current_edit_offset(current_user: bytes) -> int | None:
    """``CurrentUserAtom.offsetToCurrentEdit``, or ``None`` if malformed."""
    header = read_header(current_user, 0)
    if header is None or header.rec_type != RT_CURRENT_USER_ATOM:
        return None
    if len(current_user) < OFFSET_TO_CURRENT_EDIT + 4:
        return None
    return int(
        struct.unpack_from("<I", current_user, OFFSET_TO_CURRENT_EDIT)[0]
    )


def read_persist_directory(doc: bytes, directory: Record) -> list[tuple[int, int]]:
    """Decode runs of ``(first persist id, count)`` followed by offsets."""
    entries: list[tuple[int, int]] = []
    cursor = directory.start + HEADER_SIZE
    while cursor + 4 <= directory.end:
        info = struct.unpack_from("<I", doc, cursor)[0]
        cursor += 4
        first_id = info & PERSIST_ID_MASK
        count = info >> PERSIST_COUNT_SHIFT
        for index in range(count):
            if cursor + 4 > directory.end:
                break
            entries.append((first_id + index, struct.unpack_from("<I", doc, cursor)[0]))
            cursor += 4
    return entries


def walk_persist_chain(doc: bytes, current_edit: int) -> dict[int, int]:
    """Fold the persist directories, newest edit first, into id -> offset."""
    offsets: dict[int, int] = {}
    seen: set[int] = set()
    edit_offset = current_edit
    while edit_offset > 0 and edit_offset not in seen:
        seen.add(edit_offset)
        edit = read_header(doc, edit_offset)
        if edit is None or edit.rec_type != RT_USER_EDIT_ATOM:
            break
        payload = edit.start + HEADER_SIZE
        if payload + OFFSET_PERSIST_DIRECTORY + 4 > edit.end:
            break
        last_edit = struct.unpack_from("<I", doc, payload + OFFSET_LAST_EDIT)[0]
        directory_offset = struct.unpack_from(
            "<I", doc, payload + OFFSET_PERSIST_DIRECTORY
        )[0]
        directory = read_header(doc, directory_offset)
        if directory is not None and directory.rec_type == RT_PERSIST_DIRECTORY_ATOM:
            for persist_id, offset in read_persist_directory(doc, directory):
                # The newest edit is walked first, so the first directory
                # naming an id supplies its live offset.
                offsets.setdefault(persist_id, offset)
        edit_offset = last_edit
    return offsets


def vba_persist_id(doc: bytes, offsets: dict[int, int]) -> int | None:
    """The VBA project's persist id, from the DocumentContainer."""
    for offset in offsets.values():
        record = read_header(doc, offset)
        if record is None or record.rec_type != RT_DOCUMENT_CONTAINER:
            continue
        vba_info = find_descendant(doc, record, RT_VBA_INFO)
        if vba_info is None:
            continue
        atom = find_descendant(doc, vba_info, RT_VBA_INFO_ATOM)
        if atom is not None and atom.end - atom.start >= HEADER_SIZE + 4:
            return int(struct.unpack_from("<I", doc, atom.start + HEADER_SIZE)[0])
    return None


# ---------------------------------------------------------------------------
# Locating the storage
# ---------------------------------------------------------------------------

def decode_storage(doc: bytes, record: Record) -> bytes | None:
    """The CFB held by an ``ExOleObjStg`` record, inflated when compressed."""
    body = doc[record.start + HEADER_SIZE:record.end]
    if record.rec_instance == 0:
        return body if body[:8] == CFB_MAGIC else None
    if len(body) < 4:
        return None
    declared = struct.unpack_from("<I", body, 0)[0]
    try:
        # PowerPoint truncates the deflate stream without a proper stream
        # end, so a plain zlib.decompress() raises on the missing tail.
        # A decompressobj yields everything the declared bytes encode and
        # simply stops, which is what the tolerant finish is for.
        inflated = zlib.decompressobj().decompress(body[4:])
    except zlib.error:
        return None
    # A short inflate means the deflate stream really was damaged rather
    # than merely unterminated; the caller's scan can then try another
    # record instead of parsing a truncated CFB.
    if len(inflated) != declared or inflated[:8] != CFB_MAGIC:
        return None
    return inflated


def holds_vba_project(storage: bytes) -> bool:
    """True when these bytes parse as a CFB carrying a VBA project."""
    try:
        inner = CFB.from_bytes(storage)
        try:
            inner.get_stream_in_storage("VBA", "dir")
        except KeyError:
            inner.get_stream("dir")
    except Exception:
        return False
    return True


def scan_for_storage(doc: bytes) -> tuple[Record, bytes] | None:
    """Last ``ExOleObjStg`` record holding a VBA project, containers included.

    Used when the persist chain does not resolve.  Records are appended as
    a presentation is edited, so the last match is the newest.
    """
    found: tuple[Record, bytes] | None = None

    def walk(start: int, limit: int) -> None:
        nonlocal found
        cursor = start
        while cursor + HEADER_SIZE <= limit:
            record = read_header(doc, cursor)
            if record is None or record.end > limit:
                return
            if record.is_container:
                walk(record.start + HEADER_SIZE, record.end)
            elif record.rec_type == RT_EX_OLE_OBJ_STG:
                storage = decode_storage(doc, record)
                if storage is not None and holds_vba_project(storage):
                    found = (record, storage)
            cursor = record.end

    walk(0, len(doc))
    return found


def locate(outer: CFB) -> tuple[Record, bytes]:
    """The presentation's live VBA project record and its inflated CFB."""
    try:
        doc = outer.get_stream(DOC_STREAM)
    except KeyError:
        raise VBAProjectError(
            f"Not a binary PowerPoint presentation: no {DOC_STREAM!r} stream."
        ) from None
    try:
        current_user = outer.get_stream(CURRENT_USER_STREAM)
    except KeyError:
        current_user = b""
    current_edit = current_edit_offset(current_user) if current_user else None
    if current_edit is not None:
        offsets = walk_persist_chain(doc, current_edit)
        persist_id = vba_persist_id(doc, offsets)
        offset = offsets.get(persist_id) if persist_id is not None else None
        record = read_header(doc, offset) if offset is not None else None
        if record is not None and record.rec_type == RT_EX_OLE_OBJ_STG:
            storage = decode_storage(doc, record)
            if storage is not None and holds_vba_project(storage):
                return record, storage
    located = scan_for_storage(doc)
    if located is None:
        raise VBAProjectError(
            "Presentation contains no VBA project (no ExOleObjStg record "
            "holds one). Make sure the presentation has macros."
        )
    return located


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------

def extract_vba_storage(outer: CFB) -> bytes:
    """Return the embedded VBA project CFB of a binary ``.ppt``.

    Raises :class:`~pyopenvba.exceptions.VBAProjectError` when the
    presentation carries no VBA project.
    """
    return locate(outer)[1]


def replace_vba_storage(outer: CFB, storage: bytes) -> None:
    """Splice a modified VBA project CFB back into the presentation.

    The ``ExOleObjStg`` record is rebuilt at its current position and every
    absolute offset past it -- persist directory entries, the user-edit
    chain and the ``CurrentUserAtom`` -- is shifted by the size delta.
    ``outer`` is updated in place, ready for :meth:`CFB.to_bytes`.
    """
    record, _ = locate(outer)
    doc = outer.get_stream(DOC_STREAM)

    # Persist-referenced records are top-level, and the splice below relies
    # on that: resizing a record nested inside a container would leave the
    # parent's length wrong.  The fallback scan can surface nested records,
    # so this checks rather than assumes.
    if not any(
        top.start == record.start and top.rec_type == RT_EX_OLE_OBJ_STG
        for top in top_level_records(doc)
    ):
        raise VBAProjectError(
            "The VBA project record is nested inside a container record; "
            "rewriting it in place would corrupt the presentation."
        )

    deflated = zlib.compress(storage, 6)
    body = struct.pack("<I", len(storage)) + deflated
    header = struct.pack(
        "<HHI",
        (record.ver_instance & 0xF) | (1 << 4),  # recVer kept, recInstance 1
        RT_EX_OLE_OBJ_STG,
        len(body),
    )
    delta = HEADER_SIZE + len(body) - (record.end - record.start)
    updated = bytearray(doc[:record.start] + header + body + doc[record.end:])

    def shift(value: int) -> int:
        return value + delta if value > record.start else value

    # Record starts only move when they sit after the edit point, so walking
    # the UPDATED stream finds each carrier where it now lives while the
    # stored offset values still need the delta applied.
    for top in top_level_records(bytes(updated)):
        if top.rec_type == RT_USER_EDIT_ATOM:
            payload = top.start + HEADER_SIZE
            for field in (OFFSET_LAST_EDIT, OFFSET_PERSIST_DIRECTORY):
                at = payload + field
                if at + 4 <= top.end:
                    struct.pack_into(
                        "<I", updated, at,
                        shift(struct.unpack_from("<I", updated, at)[0]),
                    )
        elif top.rec_type == RT_PERSIST_DIRECTORY_ATOM:
            cursor = top.start + HEADER_SIZE
            while cursor + 4 <= top.end:
                info = struct.unpack_from("<I", updated, cursor)[0]
                cursor += 4
                for _ in range(info >> PERSIST_COUNT_SHIFT):
                    if cursor + 4 > top.end:
                        break
                    struct.pack_into(
                        "<I", updated, cursor,
                        shift(struct.unpack_from("<I", updated, cursor)[0]),
                    )
                    cursor += 4

    outer.write_stream(DOC_STREAM, bytes(updated))

    try:
        current_user = bytearray(outer.get_stream(CURRENT_USER_STREAM))
    except KeyError:
        return
    header_record = read_header(bytes(current_user), 0)
    if (
        header_record is not None
        and header_record.rec_type == RT_CURRENT_USER_ATOM
        and len(current_user) >= OFFSET_TO_CURRENT_EDIT + 4
    ):
        struct.pack_into(
            "<I", current_user, OFFSET_TO_CURRENT_EDIT,
            shift(
                struct.unpack_from("<I", current_user, OFFSET_TO_CURRENT_EDIT)[0]
            ),
        )
        outer.write_stream(CURRENT_USER_STREAM, bytes(current_user))
