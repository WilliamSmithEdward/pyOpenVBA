"""Attachment and multi-valued columns.

A `Complex` column does not hold its values in the row.  The row holds a
Long -- an id shared by every complex column in that row -- and the
values live in a **flat table** of their own, one row per element:

    Things.Files = 4                    the row's complex id
    f_<GUID>_Files._Files = 4           the elements that belong to it
    f_<GUID>_Files.Things_Files = 7     the element's own id, an
                                        AutoNumber unique in the table

`MSysComplexColumns` names the pairing: `ConceptualTableID` and
`ColumnName` on one side, `FlatTableID` on the other.  A multi-valued
column's flat table carries a single `Value`; an attachment's carries
`FileData`, `FileFlags`, `FileName`, `FileTimeStamp`, `FileType` and
`FileURL`.

The id is handed out per row from a counter at 0x1C of the table
definition, it is **not** reused after a delete, and a row with no
elements still takes one.

`FileData` is a container of its own::

    <u32 flag> <u32 size> <body>

`flag` 1 means the body is a zlib stream and `size` its inflated length;
0 means the body is stored as it is.  Either way the body is::

    <u32 header length, 20> <u32 1> <u32 character count>
    <extension UTF-16, NUL-terminated> <the file's own bytes>

Access decides whether to compress **by file type**, not by whether it
helps: a 72-byte PNG of nearly constant bytes is stored raw while a ZIP
of the same size would be too.  `NEVER_COMPRESSED` is the set it left
alone across 45 extensions.

Writing a compressed attachment does not reproduce Access's own bytes.
Its deflate is not zlib's -- no combination of level, memLevel, strategy
or window size reproduces a stream it wrote -- so what this writes
inflates to the same file and compresses differently.  Everything else
about the value, the framing included, is byte for byte what Access
writes, and an attachment of a type Access stores raw is byte-identical.
"""

from __future__ import annotations

import datetime as dt
import zlib
from dataclasses import dataclass, field

from pyopenvba.access_read import AccessError

#: The extensions Access stores without compressing, measured by
#: attaching one identical compressible payload under 45 of them.
NEVER_COMPRESSED = frozenset(
    {"docx", "gif", "jpeg", "jpg", "png", "pptx", "xlsx", "zip"}
)
#: Access refuses these outright as unsafe, which is a policy of its own
#: and not something this library enforces.
ACCESS_REFUSES = frozenset({"accdb", "bas", "cab", "exe", "iso", "msi"})

FILE_DATA_STORED = 0
FILE_DATA_DEFLATED = 1
#: `<u32 header length><u32 version><u32 character count>` and then the
#: extension, NUL-terminated, so the header is 12 + 2 * characters: 20
#: bytes for `txt`, 22 for `docx`, 18 for `7z`.
INNER_FIXED = 12
INNER_VERSION = 1

#: What `MSysComplexColumns.ComplexTypeObjectID` points at, by the name of
#: the `MSysComplexType_*` table it names.
ATTACHMENT_TYPE = "MSysComplexType_Attachment"


@dataclass(frozen=True)
class Attachment:
    """One file in an attachment column."""

    name: str
    data: bytes
    #: The extension without its dot.  Access derives it from `name` and
    #: stores it twice: in `FileType` and inside `FileData`.
    type: str = ""
    flags: int | None = None
    timestamp: dt.datetime | None = None

    def __post_init__(self) -> None:
        if not self.type and "." in self.name:
            object.__setattr__(self, "type", self.name.rsplit(".", 1)[1].lower())


@dataclass(frozen=True)
class ComplexColumn:
    """A complex column and the flat table holding its values."""

    table: str
    column: str
    flat_table: str
    #: `"attachment"` or the scalar type's name, e.g. `"Text"`.
    kind: str
    complex_id: int
    #: The flat table's column naming the row's complex id.
    key_column: str = field(default="")
    #: The flat table's column holding each element's own id.
    id_column: str = field(default="")

    @property
    def is_attachment(self) -> bool:
        return self.kind == "attachment"


def decode_file_data(blob: bytes) -> tuple[str, bytes]:
    """`(extension, file bytes)` from a stored `FileData` value."""
    if len(blob) < 8:
        raise AccessError("FileData is too short to hold its header")
    flag = int.from_bytes(blob[0:4], "little")
    size = int.from_bytes(blob[4:8], "little")
    if flag == FILE_DATA_DEFLATED:
        body = zlib.decompress(blob[8:])
    elif flag == FILE_DATA_STORED:
        body = blob[8:]
    else:
        raise AccessError(f"FileData carries an unknown storage flag {flag}")
    if len(body) != size:
        raise AccessError(f"FileData says {size} bytes and holds {len(body)}")
    header = int.from_bytes(body[0:4], "little")
    if header < 12 or header > len(body):
        raise AccessError(f"FileData's inner header length {header} is out of range")
    characters = int.from_bytes(body[8:12], "little")
    extension = body[12 : 12 + 2 * characters].decode("utf-16-le").rstrip("\x00")
    return extension, body[header:]


def encode_file_data(extension: str, data: bytes) -> bytes:
    """The stored form of a file, compressed the way Access would."""
    text = (extension + "\x00").encode("utf-16-le")
    body = (
        (INNER_FIXED + len(text)).to_bytes(4, "little")
        + INNER_VERSION.to_bytes(4, "little")
        + (len(extension) + 1).to_bytes(4, "little")
        + text
        + data
    )
    if extension.lower() in NEVER_COMPRESSED:
        return FILE_DATA_STORED.to_bytes(4, "little") + len(body).to_bytes(4, "little") + body
    packed = zlib.compress(body)
    return FILE_DATA_DEFLATED.to_bytes(4, "little") + len(body).to_bytes(4, "little") + packed
