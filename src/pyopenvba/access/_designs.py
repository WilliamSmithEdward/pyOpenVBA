"""Forms and reports.

Access keeps a form's or report's design in `MSysAccessStorage`, under
`Forms` or `Reports`, laid out the way a module's is under `Modules`: a
numbered folder per object with the design in a `Blob` beneath it, and a
`\\x03DirData` beside the folders listing the names.  A design folder
carries two more streams, `TypeInfo` and `PropData`, and an empty
`BlobDelta`.

The design itself is a stream of property records::

    <u32 id> <u16 code> <u32 value type> <u32 width> <u32 length>
    <length bytes>

with the ids ascending inside one object.  Three ids are not properties
but markers that open the next object, each followed by a `u16`:

* `0xFE` opens a section
* `0xFD` opens the next object at the same level
* `0xFF` opens a control, and carries a second `u16` naming its type

Ids restart at the marker, which is what tells one object's records from
the next.  Every design measured -- an empty form, a form with a label
and a text box, and a report with its three sections -- rebuilds byte for
byte from what this reads.

What the values mean is another question, and this does not pretend to
answer it: a record's `code` is the property and a handful are named in
`PROPERTY_NAMES`, the rest are handed back as they are.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from pathlib import Path

from pyopenvba.access_read import AccessError

#: `<u32 id><u16 code><u32 type><u32 width><u32 length>`.
HEAD = 18
#: Ids that open an object instead of carrying a property.
OPEN_SECTION = 0xFE
OPEN_SIBLING = 0xFD
OPEN_CONTROL = 0xFF
MARKERS = (OPEN_SECTION, OPEN_SIBLING, OPEN_CONTROL)
#: The bytes before the first record, and the four that close the stream.
DESIGN_HEADER = 10
DESIGN_TRAILER = 4

#: `MSysObjects.Type` and the navigation-pane type, per kind.  Both step
#: their object ids by one, as a macro does and unlike a module.
OBJECT_TYPES = {"form": -32768, "report": -32764}
NAV_TYPES = {"form": 32768, "report": 32764}
#: The container each kind lives under in `MSysAccessStorage`.
CONTAINERS = {"form": "Forms", "report": "Reports"}
#: And the catalog container each is filed under.
CATALOG_CONTAINERS = {"form": "Forms", "report": "Reports"}

#: Control and section types, by the `u16` an `0xFF` marker carries.
CONTROL_TYPES = {
    100: "Label",
    101: "Rectangle",
    102: "Line",
    103: "Image",
    104: "CommandButton",
    105: "OptionButton",
    106: "CheckBox",
    107: "OptionGroup",
    108: "BoundObjectFrame",
    109: "TextBox",
    110: "ListBox",
    111: "ComboBox",
    112: "Subform",
    114: "ObjectFrame",
    118: "PageBreak",
    122: "ToggleButton",
    123: "Tab",
    124: "Page",
    152: "Detail",
    155: "PageHeaderSection",
    156: "PageFooterSection",
}
#: Codes whose value is the object's name.
NAME_CODES = (20, 21)

#: Where the captured designs live, one folder stream each.
TEMPLATES = Path(__file__).parents[1] / "_templates" / "designs"
#: The record that carries a design's GUID, which the catalog row repeats.
GUID_RECORD = 208
GUID_LENGTH = 16


@dataclass(frozen=True)
class DesignRecord:
    """One property of a form, report, section or control."""

    id: int
    code: int
    value_type: int
    width: int
    value: bytes

    @property
    def name(self) -> str | None:
        return PROPERTY_NAMES.get(self.code)

    def text(self) -> str | None:
        """The value read as text, when it reads as text at all."""
        if not self.value or len(self.value) % 2:
            return None
        try:
            candidate = self.value.decode("utf-16-le").rstrip("\x00")
        except UnicodeDecodeError:
            return None
        return candidate if candidate.isprintable() else None


@dataclass(frozen=True)
class DesignObject:
    """A form or report itself, or one of its sections or controls."""

    #: The marker that opened it, or `None` for the design's own object.
    marker: int | None
    #: What the marker named: a control or section type.
    type: int | None
    code: int | None
    records: tuple[DesignRecord, ...] = field(default_factory=tuple)

    @property
    def name(self) -> str | None:
        for record in self.records:
            if record.code in NAME_CODES:
                text = record.text()
                if text:
                    return text
        return None

    @property
    def type_name(self) -> str | None:
        return CONTROL_TYPES.get(self.type) if self.type is not None else None

    @property
    def is_section(self) -> bool:
        return self.type_name in ("Detail", "PageHeaderSection", "PageFooterSection")

    def property_value(self, code: int) -> bytes | None:
        for record in self.records:
            if record.code == code:
                return record.value
        return None


@dataclass(frozen=True)
class AccessDesign:
    """A form or a report, as the file describes it."""

    name: str
    kind: str
    objects: tuple[DesignObject, ...]

    @property
    def root(self) -> DesignObject:
        return self.objects[0]

    @property
    def sections(self) -> tuple[DesignObject, ...]:
        return tuple(o for o in self.objects[1:] if o.is_section)

    @property
    def controls(self) -> tuple[DesignObject, ...]:
        """The named controls, which is what the designer shows.  A design
        also carries unnamed prototypes -- the styles new controls are cut
        from -- and those are left out."""
        return tuple(o for o in self.objects[1:] if not o.is_section and o.name)


def parse_design(blob: bytes) -> tuple[bytes, tuple[DesignObject, ...], bytes]:
    """`(header, objects, trailer)` for a design blob."""
    if len(blob) < DESIGN_HEADER + DESIGN_TRAILER:
        raise AccessError("a design blob is too short to hold its header")
    objects: list[DesignObject] = []
    records: list[DesignRecord] = []
    marker: int | None = None
    kind: int | None = None
    code: int | None = None

    def close() -> None:
        objects.append(DesignObject(marker, kind, code, tuple(records)))

    at, last = DESIGN_HEADER, -1
    while at + 6 <= len(blob):
        ident = struct.unpack_from("<I", blob, at)[0]
        if ident in MARKERS:
            close()
            records = []
            marker = ident
            code = struct.unpack_from("<H", blob, at + 4)[0]
            at += 6
            kind = code
            if ident == OPEN_CONTROL:
                kind = struct.unpack_from("<H", blob, at)[0]
                at += 2
            last = -1
            continue
        if at + HEAD > len(blob):
            break
        record_code = struct.unpack_from("<H", blob, at + 4)[0]
        value_type, width, length = struct.unpack_from("<III", blob, at + 6)
        if ident <= last or length > len(blob) - at - HEAD or value_type > 0xFFFF:
            break
        records.append(
            DesignRecord(ident, record_code, value_type, width, blob[at + HEAD : at + HEAD + length])
        )
        last = ident
        at += HEAD + length
    close()
    return blob[:DESIGN_HEADER], tuple(objects), blob[at:]


def build_design(header: bytes, objects: tuple[DesignObject, ...], trailer: bytes) -> bytes:
    """The blob for what `parse_design` gave back."""
    out = bytearray(header)
    for obj in objects:
        if obj.marker is not None:
            out += struct.pack("<IH", obj.marker, obj.code or 0)
            if obj.marker == OPEN_CONTROL:
                out += struct.pack("<H", obj.type or 0)
        for record in obj.records:
            out += struct.pack(
                "<IHIII", record.id, record.code, record.value_type, record.width, len(record.value)
            )
            out += record.value
    return bytes(out + trailer)


def template(kind: str, stream: str) -> bytes:
    """One of the captured designs an empty form or report is cut from."""
    if kind not in OBJECT_TYPES:
        raise AccessError(f"kind must be 'form' or 'report', not {kind!r}")
    path = TEMPLATES / f"{kind}.{stream}"
    if not path.exists():  # pragma: no cover - the templates ship with the package
        raise AccessError(f"the {kind} {stream} template is missing")
    return path.read_bytes()


def with_guid(blob: bytes, guid: bytes) -> bytes:
    """A design's own GUID, which its catalog row repeats.  Two objects
    sharing one is not something Access writes."""
    if len(guid) != GUID_LENGTH:
        raise AccessError(f"a design GUID is {GUID_LENGTH} bytes, not {len(guid)}")
    header, objects, trailer = parse_design(blob)
    replaced = tuple(
        DesignObject(
            obj.marker,
            obj.type,
            obj.code,
            tuple(
                DesignRecord(r.id, r.code, r.value_type, r.width, guid)
                if r.id == GUID_RECORD and len(r.value) == GUID_LENGTH
                else r
                for r in obj.records
            ),
        )
        for obj in objects
    )
    return build_design(header, replaced, trailer)


# --- naming the properties ----------------------------------------------------
# `Application.SaveAsText acForm` writes the same design with its
# properties **named** and in the same order, so walking the export and
# the blob together names the codes.  These are the ones two exports
# agreed on; a design carries hundreds more, and the rest are handed back
# as they are.
PROPERTY_CODES = {
    "AddColon": 3,
    "BorderColor": 8,
    "BorderLineStyle": 11,
    "Caption": 17,
    "Name": 20,
    "BackStyle": 29,
    "FontName": 34,
    "FontSize": 35,
    "FontWeight": 37,
    "Height": 44,
    "LabelX": 52,
    "Left": 54,
    "Top": 141,
    "Width": 150,
    "OverlapFlags": 159,
    "ForeColor": 204,
    "GUID": 376,
    "AutoHeight": 476,
    "AlternateBackColor": 572,
    "LayoutCachedLeft": 587,
    "LayoutCachedTop": 588,
    "LayoutCachedWidth": 589,
    "LayoutCachedHeight": 590,
    "ThemeFontIndex": 616,
    "BackThemeColorIndex": 617,
    "BorderThemeColorIndex": 620,
    "BorderTint": 621,
    "BorderShade": 622,
    "ForeThemeColorIndex": 623,
    "ForeTint": 624,
    "GridlineThemeColorIndex": 626,
    "GridlineShade": 628,
}
PROPERTY_NAMES = {code: name for name, code in PROPERTY_CODES.items()}
#: `Name` also appears as 21 on some objects.
PROPERTY_NAMES[21] = "Name"

# --- adding a control ---------------------------------------------------------
# A record's `id` is its slot in the object's own schema, and the schema
# differs by control type: a Label keeps its GUID at 234 and a TextBox at
# 250.  These are the slots measured on controls Access created, and they
# are what a written control has to use.
CONTROL_SLOTS: dict[str, dict[str, tuple[int, int, int, int]]] = {
    # name -> (id, code, value type, width)
    "Label": {
        "OverlapFlags": (53, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "Caption": (221, 17, 12, 4),
        "GUID": (234, 376, 9, 0),
        "LayoutCachedLeft": (282, 587, 3, 4),
        "LayoutCachedTop": (283, 588, 3, 4),
        "LayoutCachedWidth": (284, 589, 3, 4),
        "LayoutCachedHeight": (285, 590, 3, 4),
    },
    "TextBox": {
        "OverlapFlags": (55, 159, 2, 1),
        "TextAlign": (70, 379, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (250, 376, 9, 0),
        "LayoutCachedLeft": (302, 587, 3, 4),
        "LayoutCachedTop": (303, 588, 3, 4),
        "LayoutCachedWidth": (304, 589, 3, 4),
        "LayoutCachedHeight": (305, 590, 3, 4),
    },
}
#: The value Access wrote for a control it had just made.
DEFAULT_OVERLAP = 85
#: A text box carries this, and a control Access makes always has it.
DEFAULT_TEXT_ALIGN = 3
#: The `u16` an `0xFF` marker carries when it opens a trailing group of
#: controls, against 3 or 4 for the prototypes ahead of them.
CONTROLS_GROUP = 2
TYPE_CODES = {name: code for code, name in CONTROL_TYPES.items()}


def _record(slot: tuple[int, int, int, int], value: bytes) -> DesignRecord:
    ident, code, value_type, width = slot
    return DesignRecord(ident, code, value_type, width, value)


def control_object(
    control_type: str,
    name: str,
    guid: bytes,
    *,
    left: int,
    top: int,
    width: int,
    height: int,
    caption: str | None,
) -> DesignObject:
    """One control, as the records Access writes for a new one."""
    slots = CONTROL_SLOTS.get(control_type)
    if slots is None:
        raise AccessError(
            f"a {control_type} cannot be written yet; known: {', '.join(sorted(CONTROL_SLOTS))}"
        )
    values: list[tuple[str, bytes]] = [
        ("OverlapFlags", bytes((DEFAULT_OVERLAP,))),
    ]
    if control_type == "TextBox":
        values.append(("TextAlign", bytes((DEFAULT_TEXT_ALIGN,))))
    values += [
        ("Left", left.to_bytes(2, "little")),
        ("Top", top.to_bytes(2, "little")),
        ("Width", width.to_bytes(2, "little")),
        ("Height", height.to_bytes(2, "little")),
        ("Name", name.encode("utf-16-le")),
    ]
    if caption is not None:
        key = "Caption" if control_type == "Label" else "ControlSource"
        values.append((key, caption.encode("utf-16-le")))
    values += [
        ("GUID", guid),
        ("LayoutCachedLeft", left.to_bytes(4, "little")),
        ("LayoutCachedTop", top.to_bytes(4, "little")),
        ("LayoutCachedWidth", (left + width).to_bytes(4, "little")),
        ("LayoutCachedHeight", (top + height).to_bytes(4, "little")),
    ]
    records = tuple(_record(slots[key], value) for key, value in values if key in slots)
    return DesignObject(None, TYPE_CODES[control_type], None, records)


def _placed(controls: tuple[DesignObject, ...]) -> tuple[DesignObject, ...]:
    """The markers a section's controls carry, which depend on **how many
    there are**.

    One control is written as a single child, `0xFE <type>`.  Two or more
    open a group: the first `0xFF <2> <type>` and the rest `0xFD <type>`.
    Access writes it both ways -- a report whose page header holds one
    control and whose detail holds two carries both in one design -- and
    each is refused in the other's place.
    """
    out: list[DesignObject] = []
    for position, control in enumerate(controls):
        if len(controls) == 1:
            marker, code = OPEN_SECTION, control.type
        elif position == 0:
            marker, code = OPEN_CONTROL, CONTROLS_GROUP
        else:
            marker, code = OPEN_SIBLING, control.type
        out.append(DesignObject(marker, control.type, code, control.records))
    return tuple(out)


def add_control(
    blob: bytes,
    control_type: str,
    name: str,
    guid: bytes,
    *,
    section: str = "Detail",
    left: int = 0,
    top: int = 0,
    width: int = 1440,
    height: int = 240,
    caption: str | None = None,
) -> bytes:
    """A design with one more control on it.

    A control belongs to a section and is written immediately after it,
    before the section that follows.  Adding one rewrites the markers of
    the controls already there, since they depend on how many the section
    holds; see `_placed`.
    """
    header, objects, trailer = parse_design(blob)
    if any(o.name == name for o in objects):
        raise AccessError(f"this design already has an object named {name!r}")
    at = next((i for i, o in enumerate(objects) if o.is_section and o.name == section), None)
    if at is None:
        raise AccessError(f"this design has no {section!r} section")
    end = at + 1
    while end < len(objects) and not objects[end].is_section:
        end += 1
    control = control_object(
        control_type,
        name,
        guid,
        left=left,
        top=top,
        width=width,
        height=height,
        caption=caption,
    )
    placed = _placed((*objects[at + 1 : end], control))
    return build_design(header, (*objects[: at + 1], *placed, *objects[end:]), trailer)
