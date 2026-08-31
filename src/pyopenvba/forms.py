"""
UserForm designer streams: the control tree a form's code-behind does not carry.

pyOpenVBA reads and writes a project's *code*.  A UserForm's *design* --
which controls exist, how they nest, and which of their properties the
developer actually set -- lives in a separate storage inside the same
CFB, named for the form and sitting beside ``VBA/``::

    /VBA/EntryForm            the code-behind (handled by vba.py)
    /EntryForm/f              the sites: which controls, in what order
    /EntryForm/o              each control's own property record
    /EntryForm/\\x03VBFrame    text: the form's own non-default properties
    /EntryForm/i06/{f,o}      a Frame: its children live in their own storage
    /EntryForm/i06/i08/{f,o}  a Page inside a MultiPage

This module reads that tree ([MS-OFORMS]).  It does not name individual
properties -- that needs a per-class bit-to-name table this does not yet
have -- but it does expose each control's raw property mask, which is the
one thing a live host cannot tell you: MSForms writes a property into a
control's record only when it differs from that control's default, so the
mask is the set the developer chose.  A sited control read over COM
reports inherited, default and chosen values indistinguishably.

Conservative by construction.  Every structure here is length-prefixed or
counted, so a misreading collapses immediately rather than yielding a
plausible-looking control list, and this raises :class:`FormParseError`
instead of guessing.  Four independent checks have to agree:

1. the site count is consistent with ``SiteDepthsAndTypes``,
2. ``CountOfBytes`` runs exactly to the end of the ``f`` stream,
3. the per-site ``ObjectStreamSize`` values sum to exactly ``len(o)``,
4. every child storage is claimed by a site that can contain one.
"""

from __future__ import annotations

import struct
from collections.abc import Sequence
from dataclasses import dataclass, field

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

# Controls persisted as a MorphDataControl ([MS-OFORMS] 2.2.6), whose
# PropMask is 8 bytes wide.  Every other control record carries 4.
_MORPH_CLASSES = frozenset({15, 23, 24, 25, 26, 27, 28})

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


@dataclass(frozen=True)
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
    properties_set: int
    """Raw PropMask: the properties the developer set, as a bit set.

    Bit meanings are per control class and this module does not name them
    yet.  Compare two files' masks to see what differs; do not read a bit
    as a value.
    """
    property_mask_width: int
    """4 or 8 bytes.  MorphData controls carry the wider mask, so a bit
    index is only meaningful together with this."""
    children: tuple[FormControl, ...] = ()

    @property
    def is_container(self) -> bool:
        return bool(self.children) or self.clsid_cache_index in _CONTAINER_CLASSES


@dataclass(frozen=True)
class VBAForm:
    """A UserForm's designer surface."""

    name: str
    designer_source: str
    """The ``\\x03VBFrame`` text, the same block a VBE export writes."""
    properties_set: int
    """The form's own PropMask, read the same way as a control's."""
    controls: tuple[FormControl, ...] = field(default=())

    def walk(self) -> list[FormControl]:
        """Every control, depth-first, containers before their children."""
        out: list[FormControl] = []

        def visit(controls: Sequence[FormControl]) -> None:
            for control in controls:
                out.append(control)
                visit(control.children)

        visit(self.controls)
        return out


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

@dataclass(frozen=True)
class _Site:
    name: str
    id: int
    clsid_cache_index: int
    tab_index: int | None
    object_stream_size: int


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

    def read4() -> int:
        reader.align(start, 4)
        return reader.u32()

    def read2() -> int:
        reader.align(start, 2)
        return reader.u16()

    name_cb = 0
    name_compressed = False
    site_id = 0
    tab_index: int | None = None
    clsid_cache_index = 0
    object_stream_size = 0
    if mask & (1 << 0):
        packed = read4()
        name_cb = packed & 0x7FFFFFFF
        name_compressed = bool(packed & 0x80000000)
    if mask & (1 << 1):
        read4()  # TagData
    if mask & (1 << 2):
        site_id = read4()
    if mask & (1 << 3):
        read4()  # HelpContextID
    if mask & (1 << 4):
        read4()  # BitFlags
    if mask & (1 << 5):
        object_stream_size = read4()
    if mask & (1 << 6):
        tab_index = read2()
    if mask & (1 << 7):
        clsid_cache_index = read2()
    # Bit 8 carries no fixed field.  Reading two bytes for it puts every
    # name two characters late, which is the tell if you ever see one.
    if mask & (1 << 9):
        read2()  # GroupID
    if mask & (1 << 11):
        read4()  # ControlTipTextData
    if mask & (1 << 12):
        read4()  # RuntimeLicKeyData
    if mask & (1 << 13):
        read4()  # ControlSourceData
    if mask & (1 << 14):
        read4()  # RowSourceData
    reader.align(start, 4)

    # SiteExtraDataBlock opens with the name.  It is not padded before the
    # first string: aligning here shifts every name.
    name = ""
    if name_cb:
        raw = reader.take(name_cb)
        name = raw.decode(encoding if name_compressed else "utf-16-le", "replace")
    # The rest of the extra block is jumped, not walked: cbSite says where
    # the next site starts.
    reader.pos = start + 4 + cb_site
    return _Site(name, site_id, clsid_cache_index, tab_index, object_stream_size)


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
    DataBlock this reader deliberately does not walk, so both layouts are
    tried and the one whose counts prove out wins.
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


def _read_container(data: bytes, encoding: str) -> tuple[int, list[_Site]]:
    """Parse a FormControl stream into its own PropMask and its sites.

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
    mask = reader.u32()
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
    return mask, sites


# ---------------------------------------------------------------------------
# The `o` stream
# ---------------------------------------------------------------------------

def _record_mask(record: bytes, clsid_cache_index: int) -> tuple[int, int]:
    """The PropMask of one control record, with the width it was read at."""
    width = 8 if clsid_cache_index in _MORPH_CLASSES else 4
    if len(record) < 4 + width:
        raise FormParseError("control record is too short to hold its PropMask")
    version = (record[0], record[1])
    if version != _CONTROL_RECORD_VERSION:
        raise FormParseError(
            f"unsupported control record version {version[1]}.{version[0]}"
        )
    fmt = "<Q" if width == 8 else "<I"
    return int(struct.unpack_from(fmt, record, 4)[0]), width


def _morph_kind(record: bytes) -> str:
    """Type a generic MorphData from the DisplayStyle in its own record."""
    reader = _Reader(record)
    reader.skip(4)  # version and cb
    mask = reader.u32()
    reader.u32()    # the MorphData PropMask is 8 bytes; DisplayStyle is low
    for bit in (0, 1, 2, 3):  # VariousPropertyBits, BackColor, ForeColor, MaxLength
        if mask & (1 << bit):
            reader.u32()
    if mask & (1 << 4):
        reader.u8()  # BorderStyle
    if mask & (1 << 5):
        reader.u8()  # ScrollBars
    if not mask & (1 << 6):
        # Not stored means the file-format default, which is Text.
        return MORPH_DISPLAY_STYLES[1]
    return MORPH_DISPLAY_STYLES.get(reader.u8(), _UNKNOWN_CONTROL)


def _kind_of(site: _Site, record: bytes) -> str:
    if site.clsid_cache_index >= _CLASS_TABLE_BASE:
        return _ACTIVEX_CONTROL
    named = CONTROL_CLASSES.get(site.clsid_cache_index)
    if named is None:
        # An index this table does not know is a structure this reader
        # does not fully understand; the honest claim is the base surface.
        return _UNKNOWN_CONTROL
    if site.clsid_cache_index == 15:
        try:
            return _morph_kind(record)
        except FormParseError:
            return _UNKNOWN_CONTROL
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
    mask, controls = _read_level(cfb, [name], encoding)
    return VBAForm(
        name=name,
        designer_source=raw_designer.decode(encoding, "replace"),
        properties_set=mask,
        controls=controls,
    )


def read_forms(cfb: CFB, *, code_page: int = 1252) -> list[VBAForm]:
    """Read every form in a project's CFB, in directory order."""
    return [read_form(cfb, name, code_page=code_page) for name in form_names(cfb)]


def _read_level(
    cfb: CFB, path: list[str], encoding: str
) -> tuple[int, tuple[FormControl, ...]]:
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

    mask, sites = _read_container(f_stream, encoding)

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

    controls: list[FormControl] = []
    offset = 0
    for site in sites:
        record = o_stream[offset:offset + site.object_stream_size]
        offset += site.object_stream_size
        storage = substorages.pop(site.id, None)
        if storage is not None:
            child_mask, children = _read_level(cfb, [*path, storage], encoding)
            control_mask, width = child_mask, 4
        else:
            children = ()
            if record:
                control_mask, width = _record_mask(record, site.clsid_cache_index)
            else:
                control_mask, width = 0, 4
        controls.append(
            FormControl(
                name=site.name,
                kind=_kind_of(site, record),
                id=site.id,
                clsid_cache_index=site.clsid_cache_index,
                tab_index=site.tab_index,
                object_stream_size=site.object_stream_size,
                properties_set=control_mask,
                property_mask_width=width,
                children=children,
            )
        )
    if substorages:
        # A child storage no site claims means the sites were misread, or
        # this is a layout the reader does not understand.  Either way the
        # control list would be incomplete.
        orphans = ", ".join(sorted(substorages.values()))
        raise FormParseError(f"{where}: child storages claimed by no site: {orphans}")
    return mask, tuple(controls)
