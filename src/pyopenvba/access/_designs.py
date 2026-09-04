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
    "TextAlign": 379,
    "SpecialEffect": 4,
    "BorderWidth": 10,
    "ColumnWidths": 18,
    "DefaultValue": 23,
    "BackColor": 28,
    "Format": 38,
    "ValidationText": 61,
    "DecimalPlaces": 71,
    "InputMask": 72,
    "RowSource": 91,
    "StatusBarText": 135,
    "ValidationRule": 145,
    "DisplayWhen": 149,
    "ListRows": 153,
    "ListWidth": 154,
    "TextFontFamily": 243,
    "Tag": 266,
    "ControlTipText": 317,
    "OldBorderStyle": 329,
    "LeftMargin": 384,
    "TopMargin": 385,
    "LineSpacing": 386,
    "RightMargin": 388,
    "BottomMargin": 389,
    "GridlineColor": 463,
    "HoverColor": 653,
    "PressedColor": 657,
    "Gradient": 693,
    "ControlSource": 27,
    "RowSourceType": 93,
    "TabIndex": 261,
    "Picture": 0,
    "TopPadding": 700,
    "BottomPadding": 701,
    "LeftPadding": 702,
    "RightPadding": 703,
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
    "CommandButton": {
        "OverlapFlags": (49, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (105, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "Caption": (221, 17, 12, 4),
        "GUID": (245, 376, 9, 0),
        "Picture": (289, 0, 3, 4),
        "LayoutCachedLeft": (293, 587, 3, 4),
        "LayoutCachedTop": (294, 588, 3, 4),
        "LayoutCachedWidth": (295, 589, 3, 4),
        "LayoutCachedHeight": (296, 590, 3, 4),
        "BottomPadding": (356, 701, 4, 0),
        "TopPadding": (357, 700, 4, 0),
        "LeftPadding": (358, 702, 4, 0),
        "RightPadding": (359, 703, 4, 0),
    },
    "ToggleButton": {
        "OverlapFlags": (49, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (107, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "Caption": (221, 17, 12, 4),
        "GUID": (249, 376, 9, 0),
        "Picture": (291, 0, 3, 4),
        "LayoutCachedLeft": (294, 587, 3, 4),
        "LayoutCachedTop": (295, 588, 3, 4),
        "LayoutCachedWidth": (296, 589, 3, 4),
        "LayoutCachedHeight": (297, 590, 3, 4),
        "BottomPadding": (353, 701, 4, 0),
        "TopPadding": (354, 700, 4, 0),
        "LeftPadding": (355, 702, 4, 0),
        "RightPadding": (356, 703, 4, 0),
    },
    "OptionButton": {
        "OverlapFlags": (50, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (105, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (244, 376, 9, 0),
        "Picture": (290, 0, 3, 4),
        "LayoutCachedLeft": (293, 587, 3, 4),
        "LayoutCachedTop": (294, 588, 3, 4),
        "LayoutCachedWidth": (295, 589, 3, 4),
        "LayoutCachedHeight": (296, 590, 3, 4),
    },
    "CheckBox": {
        "OverlapFlags": (50, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (105, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (244, 376, 9, 0),
        "Picture": (290, 0, 3, 4),
        "LayoutCachedLeft": (293, 587, 3, 4),
        "LayoutCachedTop": (294, 588, 3, 4),
        "LayoutCachedWidth": (295, 589, 3, 4),
        "LayoutCachedHeight": (296, 590, 3, 4),
    },
    "OptionGroup": {
        "OverlapFlags": (53, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (105, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (244, 376, 9, 0),
        "Picture": (286, 0, 3, 4),
        "LayoutCachedLeft": (289, 587, 3, 4),
        "LayoutCachedTop": (290, 588, 3, 4),
        "LayoutCachedWidth": (291, 589, 3, 4),
        "LayoutCachedHeight": (292, 590, 3, 4),
    },
    "ListBox": {
        "OverlapFlags": (51, 159, 2, 1),
        "TextAlign": (65, 379, 2, 1),
        "Left": (97, 54, 3, 4),
        "Top": (98, 141, 3, 4),
        "Width": (99, 150, 3, 4),
        "Height": (100, 44, 3, 4),
        "TabIndex": (108, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "RowSourceType": (222, 93, 10, 4),
        "GUID": (248, 376, 9, 0),
        "Picture": (296, 0, 3, 4),
        "LayoutCachedLeft": (300, 587, 3, 4),
        "LayoutCachedTop": (301, 588, 3, 4),
        "LayoutCachedWidth": (302, 589, 3, 4),
        "LayoutCachedHeight": (303, 590, 3, 4),
    },
    "ComboBox": {
        "OverlapFlags": (53, 159, 2, 1),
        "TextAlign": (68, 379, 2, 1),
        "Left": (99, 54, 3, 4),
        "Top": (100, 141, 3, 4),
        "Width": (101, 150, 3, 4),
        "Height": (102, 44, 3, 4),
        "TabIndex": (110, 261, 3, 4),
        "GUID": (190, 376, 9, 0),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "RowSourceType": (222, 93, 10, 4),
        "Picture": (306, 0, 3, 4),
        "LayoutCachedLeft": (310, 587, 3, 4),
        "LayoutCachedTop": (311, 588, 3, 4),
        "LayoutCachedWidth": (312, 589, 3, 4),
        "LayoutCachedHeight": (313, 590, 3, 4),
    },
    "Rectangle": {
        "OverlapFlags": (53, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (228, 376, 9, 0),
        "LayoutCachedLeft": (281, 587, 3, 4),
        "LayoutCachedTop": (282, 588, 3, 4),
        "LayoutCachedWidth": (283, 589, 3, 4),
        "LayoutCachedHeight": (284, 590, 3, 4),
    },
    "Line": {
        "OverlapFlags": (51, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (223, 376, 9, 0),
        "LayoutCachedLeft": (276, 587, 3, 4),
        "LayoutCachedTop": (277, 588, 3, 4),
        "LayoutCachedWidth": (278, 589, 3, 4),
        "LayoutCachedHeight": (279, 590, 3, 4),
    },
    "Image": {
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (235, 376, 9, 0),
        "Picture": (282, 0, 3, 4),
        "LayoutCachedLeft": (285, 587, 3, 4),
        "LayoutCachedTop": (286, 588, 3, 4),
        "LayoutCachedWidth": (287, 589, 3, 4),
        "LayoutCachedHeight": (288, 590, 3, 4),
    },
    "PageBreak": {
        "OverlapFlags": (48, 159, 2, 1),
        "Top": (97, 141, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (223, 376, 9, 0),
    },
    "BoundObjectFrame": {
        "OverlapFlags": (53, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (245, 376, 9, 0),
        "Picture": (291, 0, 3, 4),
        "LayoutCachedLeft": (294, 587, 3, 4),
        "LayoutCachedTop": (295, 588, 3, 4),
        "LayoutCachedWidth": (296, 589, 3, 4),
        "LayoutCachedHeight": (297, 590, 3, 4),
    },
    "ObjectFrame": {
        "OverlapFlags": (55, 159, 2, 1),
        "Left": (97, 54, 3, 4),
        "Top": (98, 141, 3, 4),
        "Width": (99, 150, 3, 4),
        "Height": (100, 44, 3, 4),
        "TabIndex": (103, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (245, 376, 9, 0),
        "Picture": (285, 0, 3, 4),
        "LayoutCachedLeft": (288, 587, 3, 4),
        "LayoutCachedTop": (289, 588, 3, 4),
        "LayoutCachedWidth": (290, 589, 3, 4),
        "LayoutCachedHeight": (291, 590, 3, 4),
    },
    "Subform": {
        "OverlapFlags": (51, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (100, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (229, 376, 9, 0),
        "Picture": (276, 0, 3, 4),
        "LayoutCachedLeft": (279, 587, 3, 4),
        "LayoutCachedTop": (280, 588, 3, 4),
        "LayoutCachedWidth": (281, 589, 3, 4),
        "LayoutCachedHeight": (282, 590, 3, 4),
    },
    # A tab control has no left or top of its own: Access positions it by
    # its pages.
    "Tab": {
        "OverlapFlags": (49, 159, 2, 1),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (102, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (235, 376, 9, 0),
        "Picture": (282, 0, 3, 4),
        "LayoutCachedWidth": (288, 589, 3, 4),
        "LayoutCachedHeight": (289, 590, 3, 4),
    },
    "Page": {
        "OverlapFlags": (49, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Caption": (221, 17, 12, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (234, 376, 9, 0),
        "LayoutCachedLeft": (283, 587, 3, 4),
        "LayoutCachedTop": (284, 588, 3, 4),
        "LayoutCachedWidth": (285, 589, 3, 4),
        "LayoutCachedHeight": (286, 590, 3, 4),
        "BottomPadding": (305, 701, 4, 0),
        "TopPadding": (306, 700, 4, 0),
        "LeftPadding": (307, 702, 4, 0),
        "RightPadding": (308, 703, 4, 0),
    },
}
#: The value Access wrote for a control it had just made.
DEFAULT_OVERLAP = 85
#: What a list or combo box lists, and what Access writes for a new one.
DEFAULT_ROW_SOURCE_TYPE = "Table/Query"
#: No picture.  Access writes this on every control whose PictureData and
#: ImageData read back as -1, and on no other.
NO_PICTURE = b"\xff\xff\xff\xff"
#: What Access writes around a button it has just made, in twips.
BUTTON_PADDING = {
    "CommandButton": {"Top": 2, "Bottom": 2, "Left": 1, "Right": 1},
    "ToggleButton": {"Top": 2, "Bottom": 2, "Left": 2, "Right": 2},
    "Page": {"Top": 2, "Bottom": 2, "Left": 2, "Right": 2},
}
#: A page belongs to a tab control and nowhere else, and a tab control is
#: the only thing that holds one.
PAGE_HOLDER = "Tab"
PAGE = "Page"
#: Controls that take the focus, so a new one is given the next tab index.
#: Access omits the record when the index is 0.
TABBABLE = (
    "TextBox",
    "CommandButton",
    "ToggleButton",
    "OptionButton",
    "CheckBox",
    "OptionGroup",
    "ListBox",
    "ComboBox",
    "ObjectFrame",
    "Subform",
    "Tab",
)
#: A text box carries this, and a control Access makes always has it.
DEFAULT_TEXT_ALIGN = 3
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
    tab_index: int = 0,
) -> DesignObject:
    """One control, as the records Access writes for a new one.

    Every slot's id, code, value type and width was read back from a
    control Access itself made, and a control gets only the slots its own
    type has: a page break carries no width, an image no overlap flags, a
    combo box its GUID ahead of its name.  The records go out in id order,
    which is the order the reader expects.
    """
    slots = CONTROL_SLOTS.get(control_type)
    if slots is None:
        raise AccessError(
            f"a {control_type} cannot be written yet; known: {', '.join(sorted(CONTROL_SLOTS))}"
        )
    values: dict[str, bytes] = {
        "OverlapFlags": bytes((DEFAULT_OVERLAP,)),
        "TextAlign": bytes((DEFAULT_TEXT_ALIGN,)),
        "Left": left.to_bytes(2, "little"),
        "Top": top.to_bytes(2, "little"),
        "Width": width.to_bytes(2, "little"),
        "Height": height.to_bytes(2, "little"),
        "Name": name.encode("utf-16-le"),
        "RowSourceType": DEFAULT_ROW_SOURCE_TYPE.encode("utf-16-le"),
        "GUID": guid,
        "Picture": NO_PICTURE,
        "LayoutCachedLeft": left.to_bytes(4, "little"),
        "LayoutCachedTop": top.to_bytes(4, "little"),
        "LayoutCachedWidth": (left + width).to_bytes(4, "little"),
        "LayoutCachedHeight": (top + height).to_bytes(4, "little"),
    }
    if caption is not None:
        key = "Caption" if "Caption" in slots else "ControlSource"
        values[key] = caption.encode("utf-16-le")
    if tab_index and control_type in TABBABLE:
        values["TabIndex"] = tab_index.to_bytes(2, "little")
    for edge, twips in BUTTON_PADDING.get(control_type, {}).items():
        values[f"{edge}Padding"] = twips.to_bytes(4, "little")
    present = [(key, value) for key, value in values.items() if key in slots]
    records = tuple(
        _record(slots[key], value)
        for key, value in sorted(present, key=lambda kv: slots[kv[0]][0])
    )
    return DesignObject(None, TYPE_CODES[control_type], None, records)


def _placed(controls: tuple[DesignObject, ...]) -> tuple[DesignObject, ...]:
    """The markers a section's controls carry, which depend on **how many
    there are**.

    One control is written as a single child, `0xFE <type>`.  Two or more
    open a group: the first `0xFF <count> <type>` and the rest
    `0xFD <type>`, where `count` is how many objects the group holds, the
    opener included.  A form Access itself built with eleven controls
    carries `0xFF 11` twice -- once over the eleven prototypes and the
    detail section, once over the eleven controls -- which is what says
    the word is a count.  Get it wrong and Access does not complain: it
    opens the form and shows only as many controls as the number claims.
    """
    out: list[DesignObject] = []
    for position, control in enumerate(controls):
        if len(controls) == 1:
            marker, code = OPEN_SECTION, control.type
        elif position == 0:
            marker, code = OPEN_CONTROL, len(controls)
        else:
            marker, code = OPEN_SIBLING, control.type
        out.append(DesignObject(marker, control.type, code, control.records))
    return tuple(out)


def _children_span(objects: tuple[DesignObject, ...], at: int, stop: int) -> int:
    """How many objects directly follow `objects[at]` as its own children,
    theirs included.

    A marker says how its object was opened, so the run after a control
    belongs to that control when it opens with `0xFE` (one child) or
    `0xFF <count>` (that many), and belongs to the level above when it
    opens with `0xFD`, which marks a sibling.
    """
    first = at + 1
    if first >= stop:
        return 0
    marker = objects[first].marker
    if marker == OPEN_SECTION:
        count = 1
    elif marker == OPEN_CONTROL:
        code = objects[first].code
        count = code if isinstance(code, int) and code > 0 else 1
    else:
        return 0
    walked = first
    for _ in range(count):
        if walked >= stop:
            break
        walked += 1 + _children_span(objects, walked, stop)
    return walked - first


def _top_level(
    objects: tuple[DesignObject, ...], start: int, stop: int
) -> list[tuple[int, int]]:
    """Each control directly inside the run, as `(index, children it owns)`.

    The count a section's opening marker carries is of its own controls,
    not of everything beneath them, so a tab control's pages are stepped
    over rather than counted.
    """
    if start >= stop:
        return []
    marker = objects[start].marker
    if marker == OPEN_CONTROL:
        code = objects[start].code
        count = code if isinstance(code, int) and code > 0 else 1
    else:
        count = 1
    out: list[tuple[int, int]] = []
    at = start
    for _ in range(count):
        if at >= stop:
            break
        owned = _children_span(objects, at, stop)
        out.append((at, owned))
        at += 1 + owned
    return out


def add_control(
    blob: bytes,
    control_type: str,
    name: str,
    guid: bytes,
    *,
    section: str = "Detail",
    parent: str | None = None,
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

    `parent` names a control that holds controls of its own -- a tab
    control holding pages -- and the new one joins that group rather than
    the section's.
    """
    header, objects, trailer = parse_design(blob)
    if any(o.name == name for o in objects):
        raise AccessError(f"this design already has an object named {name!r}")
    if (control_type == PAGE) != (parent is not None):
        raise AccessError(
            f"a {PAGE} needs a parent {PAGE_HOLDER} and nothing else takes one"
        )
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
        tab_index=sum(
            1
            for o in objects
            if o.type is not None and CONTROL_TYPES.get(o.type) in TABBABLE
        ),
    )
    if parent is not None:
        return _nested(header, objects, trailer, at + 1, end, parent, control)
    owners = _top_level(objects, at + 1, end)
    kept: list[DesignObject] = []
    for index, owned in owners:
        kept.append(objects[index])
        kept.extend(objects[index + 1 : index + 1 + owned])
    # The section's own controls are re-marked; whatever each of them
    # holds rides along untouched.
    placed = _placed(tuple(objects[index] for index, _ in owners) + (control,))
    rebuilt: list[DesignObject] = []
    for position, (index, owned) in enumerate(owners):
        rebuilt.append(placed[position])
        rebuilt.extend(objects[index + 1 : index + 1 + owned])
    rebuilt.append(placed[-1])
    return build_design(header, (*objects[: at + 1], *rebuilt, *objects[end:]), trailer)


def _nested(
    header: bytes,
    objects: tuple[DesignObject, ...],
    trailer: bytes,
    start: int,
    stop: int,
    parent: str,
    control: DesignObject,
) -> bytes:
    """The design with `control` added to what `parent` holds."""
    owners = _top_level(objects, start, stop)
    found = next((pair for pair in owners if objects[pair[0]].name == parent), None)
    if found is None:
        raise AccessError(f"this design has no control named {parent!r} in that section")
    index, owned = found
    holder = objects[index]
    if holder.type is None or CONTROL_TYPES.get(holder.type) != PAGE_HOLDER:
        raise AccessError(f"{parent!r} is not a {PAGE_HOLDER}, so it holds no {PAGE}")
    children = _placed((*objects[index + 1 : index + 1 + owned], control))
    rebuilt = (
        *objects[: index + 1],
        *children,
        *objects[index + 1 + owned :],
    )
    return build_design(header, rebuilt, trailer)
