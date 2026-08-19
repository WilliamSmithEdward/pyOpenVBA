"""Declared-type decoding: the ``type_`` indirect table.

A declaration record (``DECL_BASE + var_``/``func_`` operand) stores its
declared type in one of two forms, chosen by the u16 at ``+16``:

* ``0xFFFF`` -- *plain*: ``+14`` is an OLE Automation VARTYPE byte and
  ``+15`` is a flag byte (``0x01`` ByRef, ``0x10`` user-defined-type
  member).
* anything else -- *extended*: ``+14`` is a u16 ``DECL_BASE``-relative
  offset to a **type descriptor**, whose 16-bit tag sits in the two
  bytes immediately *before* it.

Descriptor tag = ``kind | flags << 8``:

===== ==========================================================
kind  meaning and body
===== ==========================================================
0x1B  array. body ``<u32 array_info_offset><u32 element>`` where
      ``element`` is either a VARTYPE or, when its low byte is
      0x1B/0x1D/0x20, a nested descriptor tag whose body follows.
0x1D  named type reference. body ``<u16 target><u32 0x25>``;
      ``target / 8`` indexes the module type-reference table.
0x20  fixed-length string. body ``<u16 length><u32 ...>``.
===== ==========================================================

Flag byte (tag >> 8): ``0x08`` dynamic array, ``0x10`` the declaration
is a UDT member, ``0x40`` the reference names a module-local ``Type``,
``0x60`` a module-local ``Enum``; ``0x00`` on a 0x1D tag means the type
comes from a referenced type library (``Collection``, ``Worksheet``).

The **type-reference table** is a run of 10-byte entries ending
``TYPE_TABLE_GAP`` bytes before ``DECL_BASE``, preceded by a 6-byte
header whose u32 is the table's own byte length. Entry layout is
``<u16 tag><u16><u16 name_operand><u16><u16>``; the name operand
resolves through the project identifier table like any other.
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field

# Offsets inside a declaration record.
TYPE_FIELD = 14          # VARTYPE byte, or u16 descriptor offset
FLAG_FIELD = 15          # flag byte (plain form only)
DISCRIMINATOR = 16       # 0xFFFF => plain form

VARTYPE_MASK = 0x3F
FLAG_CONST = 0x40        # set in the VARTYPE byte
FLAG_BYREF = 0x01        # set in FLAG_FIELD
FLAG_MEMBER = 0x10       # set in FLAG_FIELD

# Descriptor kinds (low byte of the tag).
KIND_ARRAY = 0x1B
KIND_TYPEREF = 0x1D
KIND_FIXSTR = 0x20
KINDS = (KIND_ARRAY, KIND_TYPEREF, KIND_FIXSTR)

# Descriptor flags (high byte of the tag).
DESC_DYNAMIC = 0x08
DESC_MEMBER = 0x10
DESC_LOCAL_TYPE = 0x40
DESC_LOCAL_ENUM = 0x60

# Type-reference table geometry.
TYPE_TABLE_GAP = 26
TYPE_TABLE_STRIDE = 10
TYPE_TABLE_HEADER = 6

VARTYPE_NAMES: dict[int, str] = {
    0: "Empty", 1: "Null", 2: "Integer", 3: "Long", 4: "Single",
    5: "Double", 6: "Currency", 7: "Date", 8: "String", 9: "Object",
    10: "Error", 11: "Boolean", 12: "Variant", 13: "Unknown",
    14: "Decimal", 17: "Byte", 20: "LongLong",
}


@dataclass
class TypeRefEntry:
    """One row of the module's type-reference table."""

    tag: int
    name_operand: int
    raw: bytes


@dataclass
class DeclaredType:
    """A decoded declared type.

    ``name`` is what follows ``As`` in source. ``array`` and ``dynamic``
    describe the declaration's own shape; ``string_length`` is set for
    ``String * n``.
    """

    name: str | None = None
    array: bool = False
    dynamic: bool = False
    string_length: int | None = None
    is_const: bool = False
    is_byref: bool = False
    is_member: bool = False
    typeref_index: int | None = None
    unresolved: list[str] = field(default_factory=list)

    def render(self) -> str | None:
        """The ``As`` clause text, or None when there is nothing to say."""
        if self.string_length is not None:
            return f"String * {self.string_length}"
        return self.name


def find_type_table(module_stream: bytes, decl_base: int) -> list[TypeRefEntry]:
    """Read the module's type-reference table, or [] when absent."""
    end = decl_base - TYPE_TABLE_GAP
    if end <= TYPE_TABLE_HEADER:
        return []
    for count in range(1, 256):
        size = count * TYPE_TABLE_STRIDE
        start = end - size
        header = start - TYPE_TABLE_HEADER
        if header < 0:
            break
        if struct.unpack_from("<I", module_stream, header)[0] != size:
            continue
        rows = []
        for k in range(count):
            raw = module_stream[start + k * TYPE_TABLE_STRIDE:
                                start + (k + 1) * TYPE_TABLE_STRIDE]
            rows.append(TypeRefEntry(
                tag=struct.unpack_from("<H", raw, 0)[0],
                name_operand=struct.unpack_from("<H", raw, 4)[0],
                raw=bytes(raw),
            ))
        return rows
    return []


def _resolve_ref(index: int, table: list[TypeRefEntry], resolver) -> str | None:
    if not (0 <= index < len(table)):
        return None
    operand = table[index].name_operand
    if operand == 0xFFFF or resolver is None:
        return None
    return resolver(operand)


def _read_descriptor(module_stream: bytes, decl_base: int, offset: int,
                     table: list[TypeRefEntry], resolver,
                     out: DeclaredType, depth: int = 0) -> None:
    """Decode the descriptor at ``decl_base + offset`` into ``out``."""
    p = decl_base + offset
    if p < 2 or p + 8 > len(module_stream):
        out.unresolved.append(f"desc@{offset:#x} out of range")
        return
    tag = struct.unpack_from("<H", module_stream, p - 2)[0]
    _decode_tagged(module_stream, decl_base, p, tag, table, resolver, out, depth)


def _decode_tagged(module_stream: bytes, decl_base: int, body: int, tag: int,
                   table: list[TypeRefEntry], resolver,
                   out: DeclaredType, depth: int) -> None:
    if depth > 4 or body + 8 > len(module_stream):
        out.unresolved.append(f"tag {tag:#06x} truncated")
        return
    kind, flags = tag & 0xFF, (tag >> 8) & 0xFF
    if flags & DESC_MEMBER:
        out.is_member = True
    if kind == KIND_TYPEREF:
        index = struct.unpack_from("<H", module_stream, body)[0] // 8
        out.typeref_index = index
        out.name = _resolve_ref(index, table, resolver)
        if out.name is None:
            out.unresolved.append(f"typeref[{index}]")
    elif kind == KIND_FIXSTR:
        out.string_length = struct.unpack_from("<H", module_stream, body)[0]
        out.name = "String"
    elif kind == KIND_ARRAY:
        out.array = True
        out.dynamic = bool(flags & DESC_DYNAMIC)
        element = struct.unpack_from("<I", module_stream, body + 4)[0]
        if (element & 0xFF) in KINDS:
            # Nested descriptor: its tag is the word at body+4 and its
            # body follows immediately.
            _decode_tagged(module_stream, decl_base, body + 6, element & 0xFFFF,
                           table, resolver, out, depth + 1)
        else:
            out.name = VARTYPE_NAMES.get(element & VARTYPE_MASK)
            if out.name is None:
                out.unresolved.append(f"element vartype {element:#x}")
    else:
        out.unresolved.append(f"tag {tag:#06x}")


def read_declared_type(module_stream: bytes, decl_base: int | None,
                       record_offset: int,
                       table: list[TypeRefEntry] | None = None,
                       resolver=None) -> DeclaredType | None:
    """Decode the declared type of the record at ``decl_base + offset``.

    ``resolver`` maps a name operand to text (normally
    ``lambda op: resolve_name(op, identifiers)``).
    """
    if decl_base is None:
        return None
    p = decl_base + record_offset
    if p < 0 or p + DISCRIMINATOR + 2 > len(module_stream):
        return None
    out = DeclaredType()
    if struct.unpack_from("<H", module_stream, p + DISCRIMINATOR)[0] == 0xFFFF:
        raw = module_stream[p + TYPE_FIELD]
        out.is_const = bool(raw & FLAG_CONST)
        flags = module_stream[p + FLAG_FIELD]
        out.is_byref = bool(flags & FLAG_BYREF)
        out.is_member = bool(flags & FLAG_MEMBER)
        out.name = VARTYPE_NAMES.get(raw & VARTYPE_MASK)
        if out.name is None:
            return None
        return out
    offset = struct.unpack_from("<H", module_stream, p + TYPE_FIELD)[0]
    if offset == 0xFFFF:
        return None
    _read_descriptor(module_stream, decl_base, offset,
                     table if table is not None else [], resolver, out)
    return out
