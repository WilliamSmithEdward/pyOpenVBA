"""
UserForm designer streams: the control tree a form's code-behind does not carry.

pyOpenVBA reads and writes a project's *code*.  A UserForm's *design* --
which controls exist, how they nest, and what each one's properties are --
lives in a separate storage inside the same CFB, named for the form and
sitting beside ``VBA/``::

    /VBA/EntryForm            the code-behind (handled by vba.py)
    /EntryForm/f              the sites: which controls, in what order
    /EntryForm/o              each control's own property record
    /EntryForm/\\x03VBFrame    text: the form's own non-default properties
    /EntryForm/i06/{f,o}      a Frame: its children live in their own storage
    /EntryForm/i06/i08/{f,o}  a Page inside a MultiPage

This module reads that tree and writes it back ([MS-OFORMS]).  Property
names come from :mod:`pyopenvba._oforms_records`, which carries one table
per control class.

MSForms stores a property only when it differs from that control's
default, so the set of stored properties is the set the developer chose.
That is not something a live host can tell you: a sited control reports
inherited, default and chosen values indistinguishably.

Conservative by construction.  Every structure here is length-prefixed or
counted, so a misreading collapses immediately rather than yielding a
plausible-looking control list, and this raises :class:`FormParseError`
instead of guessing.  Four independent checks have to agree:

1. the site count is consistent with ``SiteDepthsAndTypes``,
2. ``CountOfBytes`` runs exactly to the end of the ``f`` stream,
3. the per-site ``ObjectStreamSize`` values sum to exactly ``len(o)``,
4. every child storage is claimed by a site that can contain one.

Writing is lossless: an unedited form serializes to the bytes it was read
from, which is the gate the fixtures pin.  Editing a property rewrites
only the affected control's record and patches that site's
``ObjectStreamSize`` in place -- a fixed-width field, so no length in the
``f`` stream moves.
"""

from __future__ import annotations

import struct
from collections.abc import Sequence
from dataclasses import dataclass, field

from pyopenvba._oforms_records import (
    FORM_SPEC,
    SPECS_BY_CACHE_INDEX,
    ParsedRecord,
    Size,
    parse_record,
    serialize_record,
)
from pyopenvba.cfb import CFB
from pyopenvba.exceptions import FormParseError

# [MS-OFORMS] 2.4.5 FormEmbeddedActiveXControlCached: a ClsidCacheIndex
# below 0x7FFF names a control class directly.  The VBE persists the
# specific morph indices (23/25/26/...) rather than generic MorphData,
# but 15 stays legal from other producers and resolves through
# DisplayStyle in the control's own record.
CONTROL_CLASSES: dict[int, str] = {
    7: "MSForms.Form",
    12: "MSForms.Image",
    14: "MSForms.Frame",
    15: "MSForms.MorphData",
    16: "MSForms.SpinButton",
    17: "MSForms.CommandButton",
    18: "MSForms.TabStrip",
    21: "MSForms.Label",
    23: "MSForms.TextBox",
    24: "MSForms.ListBox",
    25: "MSForms.ComboBox",
    26: "MSForms.CheckBox",
    27: "MSForms.OptionButton",
    28: "MSForms.ToggleButton",
    47: "MSForms.ScrollBar",
    57: "MSForms.MultiPage",
}

# fmDisplayStyle ([MS-OFORMS] 2.5.20.1): what a generic MorphData is.
MORPH_DISPLAY_STYLES: dict[int, str] = {
    1: "MSForms.TextBox",
    2: "MSForms.ListBox",
    3: "MSForms.ComboBox",
    4: "MSForms.CheckBox",
    5: "MSForms.OptionButton",
    6: "MSForms.ToggleButton",
    7: "MSForms.ComboBox",
}

# A ClsidCacheIndex at or above this is an index into the form's class
# table, naming an ActiveX control whose class this cannot resolve.
_CLASS_TABLE_BASE = 0x8000

# Controls that own a child storage rather than a slice of `o`.
_CONTAINER_CLASSES = frozenset({7, 14, 57})

_UNKNOWN_CONTROL = "MSForms.Control"
_ACTIVEX_CONTROL = "ActiveX.Control"

_FORM_CONTROL_VERSION = (0, 4)
_CONTROL_RECORD_VERSION = (0, 2)

# FormPropMask bits whose payload sits in FormStreamData after the
# DataBlock, and so has to be stepped over to reach the sites.
_FORM_MASK_MOUSE_ICON = 1 << 15
_FORM_MASK_FONT = 1 << 20
_FORM_MASK_PICTURE = 1 << 21

# StdFont vs TextProps ([MS-OFORMS] 2.4.6), by the first DWORD of the GUID.
_GUID_STDFONT = 0x0BE35203
_GUID_TEXTPROPS = 0xAFC20920

_VBFRAME_STREAM = "\x03VBFrame"

# Site DataBlock field widths, in mask-bit order.  Bit 8 carries no fixed
# field: reading two bytes for it puts every name two characters late.
_SITE_FIELDS: tuple[tuple[int, str, int], ...] = (
    (0, "NameData", 4),
    (1, "TagData", 4),
    (2, "ID", 4),
    (3, "HelpContextID", 4),
    (4, "BitFlags", 4),
    (5, "ObjectStreamSize", 4),
    (6, "TabIndex", 2),
    (7, "ClsidCacheIndex", 2),
    (9, "GroupID", 2),
    (11, "ControlTipTextData", 4),
    (12, "RuntimeLicKeyData", 4),
    (13, "ControlSourceData", 4),
    (14, "RowSourceData", 4),
)


@dataclass
class FormControl:
    """One control on a form, as the designer streams describe it."""

    name: str
    kind: str
    """``MSForms.TextBox`` and friends; ``ActiveX.Control`` for a control
    whose class lives in the form's class table, and ``MSForms.Control``
    when the ClsidCacheIndex is one this does not know."""
    id: int
    clsid_cache_index: int
    tab_index: int | None
    object_stream_size: int
    """Bytes this control occupies in its parent's ``o`` stream; zero for
    a container, whose own record is the ``f`` of its child storage."""
    record: ParsedRecord | None
    """The parsed property record, or ``None`` for a control whose class
    has no table here."""
    children: tuple[FormControl, ...] = ()

    @property
    def is_container(self) -> bool:
        return bool(self.children) or self.clsid_cache_index in _CONTAINER_CLASSES

    @property
    def properties_set(self) -> int:
        """Raw PropMask: the properties the developer set, as a bit set."""
        return 0 if self.record is None else self.record.mask

    @property
    def property_mask_width(self) -> int:
        """4 or 8 bytes.  MorphData controls carry the wider mask, so a
        bit index is only meaningful together with this."""
        return 8 if self.record is not None and self.record.spec.mask64 else 4

    def properties(self) -> dict[str, object]:
        """Every property this control stores, by name."""
        return {} if self.record is None else self.record.properties()

    def get(self, name: str) -> object:
        """One stored property, or ``None`` when the control does not set it."""
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Set or clear one property.  ``None`` clears it.

        A string goes to the string table, an integer to the DataBlock,
        and a :class:`~pyopenvba._oforms_records.Size` to the size field.
        Clearing removes the mask bit, which is how the control goes back
        to inheriting the default.
        """
        if self.record is None:
            raise FormParseError(
                f"{self.name!r} is a {self.kind} and has no property table here"
            )
        _set_on_record(self.record, name, value, self.name)


def _set_on_record(
    record: ParsedRecord, name: str, value: object, owner: str
) -> None:
    """Route a property edit to the right table by the value's type."""
    if isinstance(value, Size):
        record.set_size(value.width, value.height)
    elif isinstance(value, str):
        record.set_string(name, value)
    elif value is None:
        if any(f.name == name and f.kind == "str" for f in record.spec.extra):
            record.set_string(name, None)
        else:
            record.set_value(name, None)
    elif isinstance(value, bool):
        record.set_value(name, int(value))
    elif isinstance(value, int):
        record.set_value(name, value)
    else:
        raise FormParseError(
            f"{owner}: cannot store {type(value).__name__} in property {name!r}"
        )


@dataclass
class VBAForm:
    """A UserForm's designer surface."""

    name: str
    designer_source: str
    """The ``\\x03VBFrame`` text, the same block a VBE export writes."""
    controls: tuple[FormControl, ...] = ()
    _levels: list[_Level] = field(default_factory=lambda: [], repr=False)
    _encoding: str = field(default="cp1252", repr=False)

    @property
    def properties_set(self) -> int:
        """The form's own PropMask, read the same way as a control's."""
        return self._levels[0].record.mask if self._levels else 0

    def properties(self) -> dict[str, object]:
        """The form's own stored properties."""
        return self._levels[0].record.properties() if self._levels else {}

    def get(self, name: str) -> object:
        """One of the form's own properties, or ``None`` if it sets none."""
        return self.properties().get(name)

    def set_property(self, name: str, value: object) -> None:
        """Set or clear one of the form's own properties (Caption, Zoom...).

        The form's record heads its ``f`` stream rather than sitting in
        ``o``, but it is edited exactly like a control's.
        """
        if not self._levels:
            raise FormParseError(f"form {self.name!r} has no parsed record")
        _set_on_record(self._levels[0].record, name, value, self.name)

    def walk(self) -> list[FormControl]:
        """Every control, depth-first, containers before their children."""
        out: list[FormControl] = []

        def visit(controls: Sequence[FormControl]) -> None:
            for control in controls:
                out.append(control)
                visit(control.children)

        visit(self.controls)
        return out

    def control(self, name: str) -> FormControl:
        """One control by name, at any depth.  MSForms keeps names unique."""
        for control in self.walk():
            if control.name == name:
                return control
        raise KeyError(f"no control named {name!r} on form {self.name!r}")

    def write_back(self, cfb: CFB) -> None:
        """Write every edited record back into the project's CFB.

        Only the streams whose bytes actually changed are written, so an
        unedited form leaves the CFB untouched.
        """
        for level in self._levels:
            f_bytes, o_bytes = level.serialize(self._encoding)
            if o_bytes != level.o_raw:
                cfb.write_stream_at(level.path, "o", o_bytes)
                level.o_raw = o_bytes
            if f_bytes != level.f_raw:
                cfb.write_stream_at(level.path, "f", f_bytes)
                level.f_raw = f_bytes


@dataclass
class _Site:
    """One entry of a container's site array."""

    name: str
    id: int
    clsid_cache_index: int
    tab_index: int | None
    object_stream_size: int
    osz_offset: int
    """Byte offset of ObjectStreamSize inside the ``f`` stream, so a
    changed record length can be patched without moving anything."""


@dataclass
class _Level:
    """One container's streams: the form itself, a Frame, or a Page."""

    path: list[str]
    f_raw: bytes
    o_raw: bytes
    record: ParsedRecord
    form_block_len: int
    """Bytes the FormControl record occupies at the head of ``f``; editing
    the container's own properties can resize it, shifting everything
    after."""
    sites: list[_Site]
    controls: list[FormControl]

    def serialize(self, encoding: str) -> tuple[bytes, bytes]:
        """Rebuild this level's ``f`` and ``o``."""
        # The container's own record heads `f`; a resize shifts the rest,
        # which is safe because nothing inside `f` is an absolute offset --
        # cbForm is a length and cbSites is a length. Only the offsets this
        # module tracks itself have to move.
        rebuilt = serialize_record(self.record, encoding)
        delta = len(rebuilt) - self.form_block_len
        f_bytes = bytearray(
            rebuilt + self.f_raw[self.form_block_len:]
        )
        if delta:
            self.form_block_len = len(rebuilt)
            for site in self.sites:
                if site.osz_offset >= 0:
                    site.osz_offset += delta

        chunks: list[bytes] = []
        for site, control in zip(self.sites, self.controls, strict=True):
            if control.record is None or control.children:
                # A container keeps its own record in its child storage's
                # `f`, and holds no slice of `o`.
                continue
            record = serialize_record(control.record, encoding)
            chunks.append(record)
            if len(record) != site.object_stream_size:
                # ObjectStreamSize is a fixed-width DataBlock field, so
                # only its value changes; no length in `f` moves.
                if site.osz_offset < 0:
                    raise FormParseError(
                        f"{'/'.join(self.path)}: {control.name!r} grew but its "
                        "site stores no ObjectStreamSize to update"
                    )
                struct.pack_into("<I", f_bytes, site.osz_offset, len(record))
                site.object_stream_size = len(record)
                control.object_stream_size = len(record)
        return bytes(f_bytes), b"".join(chunks)


# ---------------------------------------------------------------------------
# Bounds-checked reader
# ---------------------------------------------------------------------------

class _Reader:
    """Little-endian reader that refuses to walk off the end."""

    def __init__(self, data: bytes) -> None:
        self._data = data
        self.pos = 0

    def _need(self, count: int) -> None:
        if count < 0 or self.pos + count > len(self._data):
            raise FormParseError("designer stream walked out of bounds")

    def u8(self) -> int:
        self._need(1)
        value = self._data[self.pos]
        self.pos += 1
        return value

    def u16(self) -> int:
        self._need(2)
        value = int(struct.unpack_from("<H", self._data, self.pos)[0])
        self.pos += 2
        return value

    def u32(self) -> int:
        self._need(4)
        value = int(struct.unpack_from("<I", self._data, self.pos)[0])
        self.pos += 4
        return value

    def take(self, count: int) -> bytes:
        self._need(count)
        value = self._data[self.pos:self.pos + count]
        self.pos += count
        return value

    def skip(self, count: int) -> None:
        self._need(count)
        self.pos += count

    def align(self, base: int, size: int) -> None:
        """Align to ``size`` relative to ``base`` ([MS-OFORMS] 2.1.1.2.4)."""
        over = (self.pos - base) % size
        if over:
            self.skip(size - over)


# ---------------------------------------------------------------------------
# The `f` stream
# ---------------------------------------------------------------------------

def _skip_picture(reader: _Reader) -> None:
    reader.skip(16)  # GUID
    reader.u32()     # Preamble
    reader.skip(reader.u32())


def _skip_font(reader: _Reader) -> None:
    guid = reader.take(16)
    tag = int(struct.unpack_from("<I", guid, 0)[0])
    if tag == _GUID_STDFONT:
        if reader.u8() != 0x01:
            raise FormParseError("unknown StdFont version in FormStreamData")
        reader.skip(2 + 1 + 2 + 4)  # charset, flags, weight, height
        reader.skip(reader.u8())    # face name
        return
    if tag == _GUID_TEXTPROPS:
        reader.skip(2)              # minor, major
        reader.skip(reader.u16())
        return
    raise FormParseError(f"unknown font GUID {tag:#010x} in FormStreamData")


def _read_site(reader: _Reader, encoding: str) -> _Site:
    """One OleSiteConcreteControl ([MS-OFORMS] 2.2.10.12.3).

    ``cbSite`` counts from the mask, not from the start of the record, so
    the next site begins at ``start + 4 + cbSite``.
    """
    start = reader.pos
    if reader.u16() != 0x0000:
        raise FormParseError("unsupported OleSiteConcreteControl version")
    cb_site = reader.u16()
    mask = reader.u32()

    values: dict[str, int] = {}
    osz_offset = -1
    for bit, name, size in _SITE_FIELDS:
        if not mask & (1 << bit):
            continue
        reader.align(start, size)
        if name == "ObjectStreamSize":
            osz_offset = reader.pos
        values[name] = reader.u32() if size == 4 else reader.u16()
    reader.align(start, 4)

    # SiteExtraDataBlock opens with the name.  It is not padded before the
    # first string: aligning here shifts every name.
    name = ""
    packed = values.get("NameData", 0)
    name_cb = packed & 0x7FFFFFFF
    if name_cb:
        raw = reader.take(name_cb)
        name = raw.decode(
            encoding if packed & 0x80000000 else "utf-16-le", "replace"
        )
    # The rest of the extra block is jumped, not walked: cbSite says where
    # the next site starts.
    reader.pos = start + 4 + cb_site
    return _Site(
        name=name,
        id=values.get("ID", 0),
        clsid_cache_index=values.get("ClsidCacheIndex", 0),
        tab_index=values.get("TabIndex"),
        object_stream_size=values.get("ObjectStreamSize", 0),
        osz_offset=osz_offset,
    )


def _is_trailing_record(data: bytes, offset: int) -> bool:
    """True when the bytes left after the sites are one trailing record.

    A MultiPage's ``f`` carries a MultiPage record after the FormControl
    ([MS-OFORMS] 2.2.4).  It is version-stamped and length-prefixed, so
    this checks rather than merely tolerating a remainder.
    """
    if offset + 4 > len(data):
        return False
    version = (data[offset], data[offset + 1])
    length = int(struct.unpack_from("<H", data, offset + 2)[0])
    return version == _CONTROL_RECORD_VERSION and offset + 4 + length == len(data)


def _try_site_data(
    data: bytes, start: int, encoding: str, class_table_stored: bool
) -> list[_Site] | None:
    """Read FormSiteData, or ``None`` when this layout is not the real one.

    Whether the class-table count is stored depends on a flag inside the
    DataBlock, so both layouts are tried and the one whose counts prove
    out wins.
    """
    try:
        reader = _Reader(data)
        reader.pos = start
        if class_table_stored:
            for _ in range(reader.u16()):
                if reader.u16() != 0x0000:
                    return None
                reader.skip(reader.u16())
        count_of_sites = reader.u32()
        count_of_bytes = reader.u32()
        if count_of_sites > 10_000:
            return None
        end = reader.pos + count_of_bytes
        # SiteData closes the stream, so an exact match is the cheap proof
        # that this layout is the real one.
        if end != len(data) and not _is_trailing_record(data, end):
            return None

        # SiteDepthsAndTypes: one entry per site, or one counted entry per
        # run of consecutive sites sharing a depth and type.
        depths_start = reader.pos
        accounted = 0
        while accounted < count_of_sites:
            reader.u8()  # depth
            type_or_count = reader.u8()
            if type_or_count & 0x80:
                accounted += type_or_count & 0x7F
                reader.u8()  # OptionalType
            else:
                accounted += 1
        if accounted != count_of_sites:
            return None
        reader.align(depths_start, 4)
        return [_read_site(reader, encoding) for _ in range(count_of_sites)]
    except FormParseError:
        return None


def _read_container(
    data: bytes, encoding: str
) -> tuple[ParsedRecord, list[_Site], int]:
    """Parse a FormControl stream into its own record and its sites.

    Used for the form itself and for every container control, whose child
    storage carries a FormControl of its own.
    """
    reader = _Reader(data)
    version = (reader.u8(), reader.u8())
    if version != _FORM_CONTROL_VERSION:
        raise FormParseError(
            f"not a FormControl stream (version {version[1]}.{version[0]})"
        )
    cb_form = reader.u16()
    mask = int(struct.unpack_from("<I", data, 4)[0])
    record = parse_record(data, FORM_SPEC, encoding, end=4 + cb_form)
    # The PropMask is counted inside cbForm; skip the rest of the block.
    reader.pos = 4 + cb_form
    # FormStreamData: variable-length blobs sit between the form's own data
    # and its sites, so the site array does not start at 4 + cbForm.
    if mask & _FORM_MASK_MOUSE_ICON:
        _skip_picture(reader)
    if mask & _FORM_MASK_FONT:
        _skip_font(reader)
    if mask & _FORM_MASK_PICTURE:
        _skip_picture(reader)
    sites = _try_site_data(data, reader.pos, encoding, True)
    if sites is None:
        sites = _try_site_data(data, reader.pos, encoding, False)
    if sites is None:
        raise FormParseError("FormSiteData did not reconcile")
    return record, sites, 4 + cb_form


# ---------------------------------------------------------------------------
# Typing a control
# ---------------------------------------------------------------------------

def _kind_of(site: _Site, record: ParsedRecord | None) -> str:
    if site.clsid_cache_index >= _CLASS_TABLE_BASE:
        return _ACTIVEX_CONTROL
    named = CONTROL_CLASSES.get(site.clsid_cache_index)
    if named is None:
        # An index this table does not know is a structure this reader
        # does not fully understand; the honest claim is the base surface.
        return _UNKNOWN_CONTROL
    if site.clsid_cache_index == 15 and record is not None:
        # A generic MorphData says what it is through DisplayStyle; not
        # stored means the file-format default, which is Text.
        style = record.values.get("DisplayStyle", 1)
        return MORPH_DISPLAY_STYLES.get(style, _UNKNOWN_CONTROL)
    return named


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def form_names(cfb: CFB) -> list[str]:
    """Names of the designer storages in a project's CFB.

    A form's storage sits at the root beside ``VBA/`` and holds an ``f``
    stream; that is the structural test, because the dir stream types a
    designer and a class module identically.
    """
    names: list[str] = []
    for storage in cfb.list_storages_at():
        if storage.casefold() == "vba":
            continue
        if "f" in {name.casefold() for name in cfb.list_streams_at([storage])}:
            names.append(storage)
    return names


def read_form(cfb: CFB, name: str, *, code_page: int = 1252) -> VBAForm:
    """Read one form's designer streams.

    Raises :class:`FormParseError` when the streams do not reconcile,
    rather than returning a guessed control list.
    """
    from pyopenvba.vba import encoding_for_codepage

    encoding = encoding_for_codepage(code_page)
    try:
        raw_designer = cfb.get_stream_at([name], _VBFRAME_STREAM)
    except KeyError:
        raw_designer = b""
    levels: list[_Level] = []
    controls = _read_level(cfb, [name], encoding, levels)
    return VBAForm(
        name=name,
        designer_source=raw_designer.decode(encoding, "replace"),
        controls=controls,
        _levels=levels,
        _encoding=encoding,
    )


def read_forms(cfb: CFB, *, code_page: int = 1252) -> list[VBAForm]:
    """Read every form in a project's CFB, in directory order."""
    return [read_form(cfb, name, code_page=code_page) for name in form_names(cfb)]


def _read_level(
    cfb: CFB, path: list[str], encoding: str, levels: list[_Level]
) -> tuple[FormControl, ...]:
    """Parse one container's ``f``/``o`` pair and recurse into its children."""
    where = "/".join(path)
    try:
        f_stream = cfb.get_stream_at(path, "f")
    except KeyError:
        raise FormParseError(f"{where}: no 'f' stream") from None
    try:
        o_stream = cfb.get_stream_at(path, "o")
    except KeyError:
        o_stream = b""

    form_record, sites, form_block_len = _read_container(f_stream, encoding)

    total = sum(site.object_stream_size for site in sites)
    if total != len(o_stream):
        # The strongest check available: the sites describe the whole of
        # `o` or the sites were read wrong.
        raise FormParseError(
            f"{where}: sites account for {total} bytes of 'o' but it holds "
            f"{len(o_stream)}"
        )

    # A container's own record is the `f` of its child storage, named for
    # the site id.  Matching on the id the storage name carries beats
    # rebuilding the name, whose padding this has only observed.
    substorages: dict[int, str] = {}
    for storage in cfb.list_storages_at(path):
        digits = storage[1:]
        if storage[:1].casefold() == "i" and digits.isdigit():
            substorages[int(digits)] = storage

    level = _Level(
        path=list(path),
        f_raw=f_stream,
        o_raw=o_stream,
        record=form_record,
        form_block_len=form_block_len,
        sites=sites,
        controls=[],
    )
    levels.append(level)

    offset = 0
    for site in sites:
        slice_bytes = o_stream[offset:offset + site.object_stream_size]
        offset += site.object_stream_size
        storage = substorages.pop(site.id, None)
        if storage is not None:
            children = _read_level(cfb, [*path, storage], encoding, levels)
            # The child level parsed the container's own record; find it by
            # path, since recursion appends grandchildren after it.
            record = next(
                lvl.record for lvl in levels if lvl.path == [*path, storage]
            )
        else:
            children = ()
            spec = SPECS_BY_CACHE_INDEX.get(site.clsid_cache_index)
            record = (
                parse_record(slice_bytes, spec, encoding)
                if spec is not None and slice_bytes
                else None
            )
        control = FormControl(
            name=site.name,
            kind=_kind_of(site, record),
            id=site.id,
            clsid_cache_index=site.clsid_cache_index,
            tab_index=site.tab_index,
            object_stream_size=site.object_stream_size,
            record=record,
            children=children,
        )
        level.controls.append(control)
    if substorages:
        # A child storage no site claims means the sites were misread, or
        # this is a layout the reader does not understand.  Either way the
        # control list would be incomplete.
        orphans = ", ".join(sorted(substorages.values()))
        raise FormParseError(f"{where}: child storages claimed by no site: {orphans}")
    return tuple(level.controls)
