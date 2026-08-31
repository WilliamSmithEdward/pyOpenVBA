"""
[MS-OFORMS] control records: the property table behind a control's mask.

Every control record shares one skeleton -- a four-byte header (minor
0x00, major 0x02, cbRecord), a property mask, a DataBlock of mask-selected
fields at most 4 bytes each in a fixed order with self-size alignment, an
ExtraDataBlock holding the larger values (strings, sizes), then outside
cbRecord the StreamData (pictures), TextProps, and a per-type tail.

So each control class is a *table* here and one reader walks them all.
Field order is the DataBlock order from each record's diagram, which is
also mask-bit order; mask-only booleans carry no DataBlock field and are
simply absent.

Lossless by construction, which is what makes writing safe: alignment
padding is captured and replayed (its bytes are undefined, so recomputing
them would not be byte-identical), string bytes are kept raw beside their
decoded text, pictures are kept as opaque runs, and any tail these tables
do not model is preserved verbatim.  An unmodified parse therefore
serializes to the identical bytes -- the gate the fixtures pin.

Private module: the public surface is :mod:`pyopenvba.forms`.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from typing import Literal

from pyopenvba.exceptions import FormParseError

FieldKind = Literal["u", "i", "lenBytes", "marker"]
"""``u``/``i`` are unsigned/signed integers of the field's size,
``lenBytes`` is a CountOfBytesWithCompressionFlag naming a string in the
ExtraDataBlock, and ``marker`` is the 0xFFFF placeholder for a picture
that lives in StreamData."""

ExtraKind = Literal["size8", "pos8", "str", "arrRaw"]


@dataclass(frozen=True)
class DataField:
    bit: int
    name: str
    size: Literal[1, 2, 4]
    kind: FieldKind = "u"


@dataclass(frozen=True)
class ExtraField:
    bit: int
    name: str
    kind: ExtraKind
    size_from: str | None = None


@dataclass(frozen=True)
class RecordSpec:
    type: str
    data: tuple[DataField, ...]
    extra: tuple[ExtraField, ...] = ()
    stream: tuple[tuple[int, str], ...] = ()
    """StreamData members in order; each present when its mask bit is set."""
    major: int = 0x02
    mask64: bool = False
    stop_after_extra: bool = False
    """Stop at the cb boundary and leave the rest to the caller."""
    text_props: bool = False
    raw_tail: bool = False
    """Bytes after TextProps (rgColumnInfo, TabStripTabFlags)."""


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

TEXT_PROPS_SPEC = RecordSpec(
    type="TextProps",
    # TextProps ends at its cb boundary; whatever follows (a TabStrip's tab
    # flags, MorphData's column info) belongs to the record that embeds it.
    stop_after_extra=True,
    data=(
        DataField(0, "FontName", 4, "lenBytes"),
        DataField(1, "FontEffects", 4),
        DataField(2, "FontHeight", 4),
        DataField(4, "FontCharSet", 1),
        DataField(5, "FontPitchAndFamily", 1),
        DataField(6, "ParagraphAlign", 1),
        DataField(7, "FontWeight", 2),
    ),
    extra=(ExtraField(0, "FontName", "str"),),
)

MORPH_DATA_SPEC = RecordSpec(
    type="MorphData",
    mask64=True,
    data=(
        DataField(0, "VariousPropertyBits", 4),
        DataField(1, "BackColor", 4),
        DataField(2, "ForeColor", 4),
        DataField(3, "MaxLength", 4),
        DataField(4, "BorderStyle", 1),
        DataField(5, "ScrollBars", 1),
        DataField(6, "DisplayStyle", 1),
        DataField(7, "MousePointer", 1),
        DataField(9, "PasswordChar", 2),
        DataField(10, "ListWidth", 4),
        DataField(11, "BoundColumn", 2),
        DataField(12, "TextColumn", 2, "i"),
        DataField(13, "ColumnCount", 2, "i"),
        DataField(14, "ListRows", 2),
        DataField(15, "cColumnInfo", 2),
        DataField(16, "MatchEntry", 1),
        DataField(17, "ListStyle", 1),
        DataField(18, "ShowDropButtonWhen", 1),
        DataField(20, "DropButtonStyle", 1),
        DataField(21, "MultiSelect", 1),
        DataField(22, "Value", 4, "lenBytes"),
        DataField(23, "Caption", 4, "lenBytes"),
        DataField(24, "PicturePosition", 4),
        DataField(25, "BorderColor", 4),
        DataField(26, "SpecialEffect", 4),
        DataField(27, "MouseIcon", 2, "marker"),
        DataField(28, "Picture", 2, "marker"),
        DataField(29, "Accelerator", 2),
        DataField(32, "GroupName", 4, "lenBytes"),
    ),
    extra=(
        ExtraField(8, "Size", "size8"),
        ExtraField(22, "Value", "str"),
        ExtraField(23, "Caption", "str"),
        ExtraField(32, "GroupName", "str"),
    ),
    stream=((27, "MouseIcon"), (28, "Picture")),
    text_props=True,
    raw_tail=True,  # rgColumnInfo for ComboBox / ListBox
)

COMMAND_BUTTON_SPEC = RecordSpec(
    type="CommandButton",
    data=(
        DataField(0, "ForeColor", 4),
        DataField(1, "BackColor", 4),
        DataField(2, "VariousPropertyBits", 4),
        DataField(3, "Caption", 4, "lenBytes"),
        DataField(4, "PicturePosition", 4),
        DataField(6, "MousePointer", 1),
        DataField(7, "Picture", 2, "marker"),
        DataField(8, "Accelerator", 2),
        DataField(10, "MouseIcon", 2, "marker"),
    ),
    extra=(ExtraField(3, "Caption", "str"), ExtraField(5, "Size", "size8")),
    stream=((7, "Picture"), (10, "MouseIcon")),
    text_props=True,
)

LABEL_SPEC = RecordSpec(
    type="Label",
    data=(
        DataField(0, "ForeColor", 4),
        DataField(1, "BackColor", 4),
        DataField(2, "VariousPropertyBits", 4),
        DataField(3, "Caption", 4, "lenBytes"),
        DataField(4, "PicturePosition", 4),
        DataField(6, "MousePointer", 1),
        DataField(7, "BorderColor", 4),
        DataField(8, "BorderStyle", 2),
        DataField(9, "SpecialEffect", 2),
        DataField(10, "Picture", 2, "marker"),
        DataField(11, "Accelerator", 2),
        DataField(12, "MouseIcon", 2, "marker"),
    ),
    extra=(ExtraField(3, "Caption", "str"), ExtraField(5, "Size", "size8")),
    stream=((10, "Picture"), (12, "MouseIcon")),
    text_props=True,
)

IMAGE_SPEC = RecordSpec(
    type="Image",
    data=(
        DataField(3, "BorderColor", 4),
        DataField(4, "BackColor", 4),
        DataField(5, "BorderStyle", 1),
        DataField(6, "MousePointer", 1),
        DataField(7, "PictureSizeMode", 1),
        DataField(8, "SpecialEffect", 1),
        DataField(10, "Picture", 2, "marker"),
        DataField(11, "PictureAlignment", 1),
        DataField(13, "VariousPropertyBits", 4),
        DataField(14, "MouseIcon", 2, "marker"),
    ),
    extra=(ExtraField(9, "Size", "size8"),),
    stream=((10, "Picture"), (14, "MouseIcon")),
    # Image carries no TextProps ([MS-OFORMS] 2.3.1 applies-to list).
)

SPIN_BUTTON_SPEC = RecordSpec(
    type="SpinButton",
    data=(
        DataField(0, "ForeColor", 4),
        DataField(1, "BackColor", 4),
        DataField(2, "VariousPropertyBits", 4),
        DataField(5, "Min", 4, "i"),
        DataField(6, "Max", 4, "i"),
        DataField(7, "Position", 4, "i"),
        DataField(8, "PrevEnabled", 4, "i"),
        DataField(9, "NextEnabled", 4, "i"),
        DataField(10, "SmallChange", 4, "i"),
        DataField(11, "Orientation", 4, "i"),
        DataField(12, "Delay", 4),
        DataField(13, "MouseIcon", 2, "marker"),
        DataField(14, "MousePointer", 1),
    ),
    extra=(ExtraField(3, "Size", "size8"),),
    stream=((13, "MouseIcon"),),
    # SpinButton carries no TextProps.
)

SCROLL_BAR_SPEC = RecordSpec(
    type="ScrollBar",
    data=(
        DataField(0, "ForeColor", 4),
        DataField(1, "BackColor", 4),
        DataField(2, "VariousPropertyBits", 4),
        DataField(4, "MousePointer", 1),
        DataField(5, "Min", 4, "i"),
        DataField(6, "Max", 4, "i"),
        DataField(7, "Position", 4, "i"),
        DataField(9, "PrevEnabled", 4, "i"),
        DataField(10, "NextEnabled", 4, "i"),
        DataField(11, "SmallChange", 4, "i"),
        DataField(12, "LargeChange", 4, "i"),
        DataField(13, "Orientation", 4, "i"),
        DataField(14, "ProportionalThumb", 2, "i"),
        DataField(15, "Delay", 4),
        DataField(16, "MouseIcon", 2, "marker"),
    ),
    extra=(ExtraField(3, "Size", "size8"),),
    stream=((16, "MouseIcon"),),
    # ScrollBar carries no TextProps.
)

TAB_STRIP_SPEC = RecordSpec(
    type="TabStrip",
    data=(
        DataField(0, "ListIndex", 4, "i"),
        DataField(1, "BackColor", 4),
        DataField(2, "ForeColor", 4),
        DataField(5, "ItemsSize", 4),
        DataField(6, "MousePointer", 1),
        DataField(8, "TabOrientation", 4),
        DataField(9, "TabStyle", 4),
        DataField(11, "TabFixedWidth", 4),
        DataField(12, "TabFixedHeight", 4),
        DataField(15, "TipStringsSize", 4),
        DataField(17, "NamesSize", 4),
        DataField(18, "VariousPropertyBits", 4),
        DataField(20, "TabsAllocated", 4),
        DataField(21, "TagsSize", 4),
        DataField(22, "TabData", 4),
        DataField(23, "AcceleratorsSize", 4),
        DataField(24, "MouseIcon", 2, "marker"),
    ),
    extra=(
        ExtraField(4, "Size", "size8"),
        ExtraField(5, "Items", "arrRaw", "ItemsSize"),
        ExtraField(15, "TipStrings", "arrRaw", "TipStringsSize"),
        ExtraField(17, "TabNames", "arrRaw", "NamesSize"),
        ExtraField(21, "Tags", "arrRaw", "TagsSize"),
        ExtraField(23, "Accelerators", "arrRaw", "AcceleratorsSize"),
    ),
    stream=((24, "MouseIcon"),),
    text_props=True,
    raw_tail=True,  # TabStripTabFlags
)

FORM_SPEC = RecordSpec(
    type="Form",
    major=0x04,
    # The sites follow the form's own block; forms.py walks those.
    stop_after_extra=True,
    data=(
        DataField(1, "BackColor", 4),
        DataField(2, "ForeColor", 4),
        DataField(3, "NextAvailableID", 4),
        DataField(6, "BooleanProperties", 4),
        DataField(7, "BorderStyle", 1),
        DataField(8, "MousePointer", 1),
        DataField(9, "ScrollBars", 1),
        DataField(13, "GroupCnt", 4, "i"),
        DataField(15, "MouseIcon", 2, "marker"),
        DataField(16, "Cycle", 1),
        DataField(17, "SpecialEffect", 1),
        DataField(18, "BorderColor", 4),
        DataField(19, "Caption", 4, "lenBytes"),
        DataField(20, "Font", 2, "marker"),
        DataField(21, "Picture", 2, "marker"),
        DataField(22, "Zoom", 4),
        DataField(23, "PictureAlignment", 1),
        DataField(25, "PictureSizeMode", 1),
        DataField(26, "ShapeCookie", 4),
        DataField(27, "DrawBuffer", 4),
    ),
    extra=(
        ExtraField(10, "DisplayedSize", "size8"),
        ExtraField(11, "LogicalSize", "size8"),
        ExtraField(12, "ScrollPosition", "pos8"),
        ExtraField(19, "Caption", "str"),
    ),
)

SPECS_BY_CACHE_INDEX: dict[int, RecordSpec] = {
    12: IMAGE_SPEC,
    15: MORPH_DATA_SPEC,
    16: SPIN_BUTTON_SPEC,
    17: COMMAND_BUTTON_SPEC,
    18: TAB_STRIP_SPEC,
    21: LABEL_SPEC,
    23: MORPH_DATA_SPEC,  # TextBox
    24: MORPH_DATA_SPEC,  # ListBox
    25: MORPH_DATA_SPEC,  # ComboBox
    26: MORPH_DATA_SPEC,  # CheckBox
    27: MORPH_DATA_SPEC,  # OptionButton
    28: MORPH_DATA_SPEC,  # ToggleButton
    47: SCROLL_BAR_SPEC,
}




# ---------------------------------------------------------------------------
# Parsed values
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Size:
    """A pair from an fmSize or fmPosition, in HIMETRIC units."""

    width: int
    height: int


@dataclass
class StoredString:
    """A string property, kept both decoded and as its stored bytes.

    The raw bytes are replayed verbatim unless the text was edited, so an
    untouched record serializes byte for byte -- an MBCS encoder is not
    guaranteed to reproduce the exact bytes it decoded.
    """

    text: str
    compressed: bool
    raw: bytes = b""
    edited: bool = False


@dataclass
class ParsedRecord:
    """One control's record: its mask, its values, and everything opaque.

    Mutable on purpose.  Writing a form means editing this and serializing
    it back, so every field the tables do not model is carried here and
    replayed, which is what keeps an untouched round-trip byte-exact.
    """

    spec: RecordSpec
    mask: int
    values: dict[str, int] = field(default_factory=lambda: {})
    strings: dict[str, StoredString] = field(default_factory=lambda: {})
    sizes: dict[str, Size] = field(default_factory=lambda: {})
    arrays: dict[str, bytes] = field(default_factory=lambda: {})
    pictures: dict[str, bytes] = field(default_factory=lambda: {})
    pads: dict[str, bytes] = field(default_factory=lambda: {})
    """Alignment gaps, captured because the spec leaves their bytes
    undefined; writing them back as zeros would not be byte-identical."""
    text_props: ParsedRecord | None = None
    tail_raw: bytes = b""

    def has(self, name: str) -> bool:
        """True when the record actually stores this property."""
        bit = self._bit_of(name)
        return bit is not None and bool(self.mask & (1 << bit))

    def _bit_of(self, name: str) -> int | None:
        for data_field in self.spec.data:
            if data_field.name == name:
                return data_field.bit
        for extra in self.spec.extra:
            if extra.name == name:
                return extra.bit
        return None

    def properties(self) -> dict[str, object]:
        """Every property this record stores, by name.

        A picture is reported by its byte length rather than the 0xFFFF
        placeholder the DataBlock holds, which carries no information.
        """
        out: dict[str, object] = {}
        for name, value in self.values.items():
            if name in self.strings:
                out[name] = self.strings[name].text
            elif name in self.pictures:
                out[name] = f"<{len(self.pictures[name])} bytes>"
            else:
                out[name] = value
        out.update(self.sizes)
        for name, blob in self.arrays.items():
            out[name] = f"<{len(blob)} bytes>"
        if self.text_props is not None:
            for name, value in self.text_props.properties().items():
                out[f"Font.{name}"] = value
        return out

    # -- mutation ---------------------------------------------------------

    def set_value(self, name: str, value: int | None) -> None:
        """Set or clear a numeric DataBlock field, adjusting the mask."""
        spec_field = next(
            (f for f in self.spec.data if f.name == name and f.kind != "marker"),
            None,
        )
        if spec_field is None:
            raise FormParseError(f"{self.spec.type} has no numeric field {name!r}")
        self._set_bit(spec_field.bit, value is not None)
        if value is None:
            self.values.pop(name, None)
        else:
            self.values[name] = value

    def set_string(self, name: str, text: str | None) -> None:
        """Set or clear a string property (Caption, Value, GroupName...)."""
        extra = next(
            (f for f in self.spec.extra if f.name == name and f.kind == "str"), None
        )
        if extra is None:
            raise FormParseError(f"{self.spec.type} has no string property {name!r}")
        if text is None:
            self._set_bit(extra.bit, False)
            self.strings.pop(name, None)
            self.values.pop(name, None)
            return
        existing = self.strings.get(name)
        # A new string is compressed when every character fits one byte,
        # the same choice the VBE makes; an existing one keeps its stored
        # compression so nothing else in the record shifts.
        compressed = (
            existing.compressed if existing else all(ord(c) <= 0xFF for c in text)
        )
        self._set_bit(extra.bit, True)
        self.strings[name] = StoredString(text, compressed, b"", edited=True)

    def set_size(self, width: int, height: int) -> None:
        """Set the control's size, in HIMETRIC units."""
        extra = next((f for f in self.spec.extra if f.kind == "size8"), None)
        if extra is None:
            raise FormParseError(f"{self.spec.type} has no Size")
        self._set_bit(extra.bit, True)
        self.sizes[extra.name] = Size(width, height)

    def _set_bit(self, bit: int, on: bool) -> None:
        if on:
            self.mask |= 1 << bit
        else:
            self.mask &= ~(1 << bit)


class _Reader:
    """Bounds-checked little-endian reader over one control record."""

    def __init__(self, data: bytes, pos: int = 0) -> None:
        self._data = data
        self.pos = pos

    def _need(self, count: int) -> None:
        if count < 0 or self.pos + count > len(self._data):
            raise FormParseError(
                f"control record read out of bounds at {self.pos}+{count}"
                f"/{len(self._data)}"
            )

    def integer(self, size: int, signed: bool) -> int:
        self._need(size)
        fmt = _SIGNED_FMT[size] if signed else _UNSIGNED_FMT[size]
        value = int(struct.unpack_from("<" + fmt, self._data, self.pos)[0])
        self.pos += size
        return value

    def take(self, count: int) -> bytes:
        self._need(count)
        value = self._data[self.pos:self.pos + count]
        self.pos += count
        return value

    def pad(self, base: int, size: int) -> bytes:
        over = (self.pos - base) % size
        return self.take(size - over) if over else b""


_SIGNED_FMT = {1: "b", 2: "h", 4: "i"}
_UNSIGNED_FMT = {1: "B", 2: "H", 4: "I"}


def _pack(value: int, size: int, signed: bool) -> bytes:
    if not signed:
        value &= (1 << (size * 8)) - 1
    return struct.pack(
        "<" + (_SIGNED_FMT[size] if signed else _UNSIGNED_FMT[size]), value
    )


def _gap(body: bytearray, header_size: int, size: int, captured: bytes) -> bytes:
    """Alignment bytes for the next field, replaying what was captured."""
    over = (len(body) + header_size) % size
    if not over:
        return b""
    needed = size - over
    return captured if len(captured) == needed else bytes(needed)


# ---------------------------------------------------------------------------
# Parse and serialize
# ---------------------------------------------------------------------------

def parse_record(
    data: bytes,
    spec: RecordSpec,
    encoding: str,
    *,
    pos: int = 0,
    end: int | None = None,
) -> ParsedRecord:
    """Parse one control record, refusing anything the table cannot explain.

    ``end`` bounds the trailing blocks -- a site's ``ObjectStreamSize``,
    or the stream's length.
    """
    limit = len(data) if end is None else end
    reader = _Reader(data, pos)
    base = reader.pos
    minor = reader.integer(1, False)
    major = reader.integer(1, False)
    if (minor, major) != (0x00, spec.major):
        raise FormParseError(
            f"{spec.type}: not a control record (version {major}.{minor})"
        )
    cb = reader.integer(2, False)
    mask = reader.integer(4, False)
    if spec.mask64:
        mask |= reader.integer(4, False) << 32
    record = ParsedRecord(spec, mask)

    # DataBlock: fields in table order, each aligned to its own size.
    for spec_field in spec.data:
        if not record.has(spec_field.name):
            continue
        gap = reader.pad(base, spec_field.size)
        if gap:
            record.pads[f"before:{spec_field.name}"] = gap
        record.values[spec_field.name] = reader.integer(
            spec_field.size, spec_field.kind == "i"
        )
    gap = reader.pad(base, 4)
    if gap:
        record.pads["data:end"] = gap

    # ExtraDataBlock, in table order; strings pad to 4 after their bytes.
    for extra in spec.extra:
        if not mask & (1 << extra.bit):
            continue
        if extra.kind in ("size8", "pos8"):
            record.sizes[extra.name] = Size(
                reader.integer(4, True), reader.integer(4, True)
            )
        elif extra.kind == "str":
            packed = record.values.get(extra.name, 0)
            raw = reader.take(packed & 0x7FFFFFFF)
            compressed = bool(packed & 0x80000000)
            record.strings[extra.name] = StoredString(
                raw.decode(encoding if compressed else "utf-16-le", "replace"),
                compressed,
                raw,
            )
            gap = reader.pad(base, 4)
            if gap:
                record.pads[f"str:{extra.name}"] = gap
        else:
            record.arrays[extra.name] = reader.take(
                record.values.get(extra.size_from or "", 0)
            )

    after_extra = base + 4 + cb
    if spec.stop_after_extra:
        if reader.pos != after_extra:
            raise FormParseError(
                f"{spec.type}: cb mismatch ({reader.pos - base - 4} != {cb})"
            )
        return record
    if reader.pos > after_extra:
        raise FormParseError(
            f"{spec.type}: record overran cb by {reader.pos - after_extra}"
        )
    if reader.pos < after_extra:
        # Bytes inside cb these tables do not account for would be dropped
        # on write; refuse rather than corrupt the form silently.
        raise FormParseError(
            f"{spec.type}: {after_extra - reader.pos} unmodelled bytes inside cb"
        )

    for bit, name in spec.stream:
        if mask & (1 << bit):
            record.pictures[name] = _read_guid_and_picture(reader)

    if spec.text_props:
        nested_end = reader.pos + _record_length(data, reader.pos)
        record.text_props = parse_record(
            data, TEXT_PROPS_SPEC, encoding, pos=reader.pos, end=nested_end
        )
        reader.pos = nested_end

    if reader.pos < limit:
        if not spec.raw_tail:
            raise FormParseError(
                f"{spec.type}: {limit - reader.pos} unexpected trailing bytes"
            )
        record.tail_raw = reader.take(limit - reader.pos)
    return record


def serialize_record(record: ParsedRecord, encoding: str) -> bytes:
    """Rebuild a record's bytes.  An unedited record round-trips exactly."""
    spec = record.spec

    # Rebuild the string length fields first: an edited string changes its
    # CountOfBytesWithCompressionFlag, which the DataBlock carries.
    for extra in spec.extra:
        if extra.kind != "str" or not record.has(extra.name):
            continue
        stored = record.strings.get(extra.name)
        if stored is None:
            continue
        raw = _string_bytes(stored, encoding)
        record.values[extra.name] = (len(raw) & 0x7FFFFFFF) | (
            0x80000000 if stored.compressed else 0
        )

    # The DataBlock aligns against the record base, so the header counts.
    header_size = 4 + (8 if spec.mask64 else 4)
    body = bytearray()
    for spec_field in spec.data:
        if not record.has(spec_field.name):
            continue
        body += _gap(
            body,
            header_size,
            spec_field.size,
            record.pads.get(f"before:{spec_field.name}", b""),
        )
        body += _pack(
            record.values.get(spec_field.name, 0),
            spec_field.size,
            spec_field.kind == "i",
        )
    body += _gap(body, header_size, 4, record.pads.get("data:end", b""))

    for extra in spec.extra:
        if not record.mask & (1 << extra.bit):
            continue
        if extra.kind in ("size8", "pos8"):
            size = record.sizes.get(extra.name, Size(0, 0))
            body += _pack(size.width, 4, True) + _pack(size.height, 4, True)
        elif extra.kind == "str":
            stored = record.strings.get(extra.name)
            body += b"" if stored is None else _string_bytes(stored, encoding)
            body += _gap(
                body, header_size, 4, record.pads.get(f"str:{extra.name}", b"")
            )
        else:
            body += record.arrays.get(extra.name, b"")

    cb = len(body) + (8 if spec.mask64 else 4)
    if cb > 0xFFFF:
        # cb is a u16.  Letting it wrap writes a record the reader then
        # overruns, so this refuses before any byte lands.
        raise FormParseError(
            f"{spec.type}: too much data for one record ({cb} bytes; the "
            "format caps a record at 65535) -- a caption or other text is "
            "too long"
        )
    out = bytearray(struct.pack("<BBH", 0x00, spec.major, cb))
    out += _pack(record.mask & 0xFFFFFFFF, 4, False)
    if spec.mask64:
        out += _pack((record.mask >> 32) & 0xFFFFFFFF, 4, False)
    out += body

    for bit, name in spec.stream:
        if not record.mask & (1 << bit):
            continue
        picture = record.pictures.get(name)
        if picture is None:
            raise FormParseError(f"{spec.type}: masked StreamData {name} has no bytes")
        out += picture
    if spec.text_props and record.text_props is not None:
        out += serialize_record(record.text_props, encoding)
    out += record.tail_raw
    return bytes(out)


def _string_bytes(stored: StoredString, encoding: str) -> bytes:
    if not stored.edited:
        return stored.raw
    return stored.text.encode(encoding if stored.compressed else "utf-16-le", "replace")


def _record_length(data: bytes, pos: int) -> int:
    """Header plus cb: how far a nested record reaches."""
    if pos + 4 > len(data):
        raise FormParseError("nested record header runs past the end")
    return 4 + int(struct.unpack_from("<H", data, pos + 2)[0])


def _read_guid_and_picture(reader: _Reader) -> bytes:
    """A GuidAndPicture: 16-byte CLSID, 4-byte preamble, 4-byte size, data."""
    start = reader.pos
    reader.take(16)
    reader.integer(4, False)  # preamble
    size = reader.integer(4, False)
    reader.take(size)
    total = reader.pos - start
    reader.pos = start
    return reader.take(total)
