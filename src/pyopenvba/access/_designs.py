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
    119: "CustomControl",
    123: "Tab",
    124: "Page",
    126: "Attachment",
    128: "WebBrowser",
    129: "NavigationControl",
    133: "Chart",
    134: "EdgeBrowser",
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
        "GUID": (249, 376, 9, 0),
        "Caption": (231, 17, 12, 4),
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
        # A subform has no ControlSource; what it shows is its
        # SourceObject, which is not written here.
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
        "Name": (220, 20, 10, 4),
        "Caption": (232, 17, 12, 4),
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
    "CustomControl": {
        "OverlapFlags": (54, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (101, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (240, 376, 9, 0),
        "Picture": (284, 0, 3, 4),
        "LayoutCachedLeft": (287, 587, 3, 4),
        "LayoutCachedTop": (288, 588, 3, 4),
        "LayoutCachedWidth": (289, 589, 3, 4),
        "LayoutCachedHeight": (290, 590, 3, 4),
    },
    # An attachment control names its column through ControlSource, and
    # Access writes its tab index after the GUID rather than before.
    "Attachment": {
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (220, 20, 10, 4),
        "ControlSource": (221, 27, 12, 4),
        "GUID": (243, 376, 9, 0),
        "TabIndex": (307, 261, 3, 4),
        "Picture": (308, 0, 3, 4),
        "LayoutCachedLeft": (314, 587, 3, 4),
        "LayoutCachedTop": (315, 588, 3, 4),
        "LayoutCachedWidth": (316, 589, 3, 4),
        "LayoutCachedHeight": (317, 590, 3, 4),
    },
    "WebBrowser": {
        "OverlapFlags": (54, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (101, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (239, 376, 9, 0),
        "Picture": (289, 0, 3, 4),
        "LayoutCachedLeft": (292, 587, 3, 4),
        "LayoutCachedTop": (293, 588, 3, 4),
        "LayoutCachedWidth": (294, 589, 3, 4),
        "LayoutCachedHeight": (295, 590, 3, 4),
    },
}

#: The design's own object, which has no control type of its own.
DESIGN_OBJECT = "_Design"
#: Value types whose length comes from the value rather than the slot.
VARIABLE_VALUE_TYPES = (10, 11, 12)
#: Both string types, which differ in what Access lets you put in them --
#: 12 takes an expression -- and not in how the bytes are written.
TEXT_VALUE_TYPES = (10, 12)

#: Where each named property sits in each object type's schema, as
#: `(id, code, value type, width, fixed length)`.
#:
#: A record's id is its slot and the slot differs by object type, so
#: changing a property an object does not already carry means knowing the
#: id Access would have given it.  Every entry was read off an object
#: Access itself wrote; across five databases every id, code and value
#: type agreed, and so did every length except those of the strings,
#: whose length is their text's.
PROPERTY_SLOTS: dict[str, dict[str, tuple[int, int, int, int, int]]] = {
    "Attachment": {
        "BackStyle": (50, 29, 2, 1, 1),
        "ControlSource": (221, 27, 12, 4, 0),
        "BorderLineStyle": (53, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (243, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "LabelX": (302, 52, 3, 4, 4),
        "AddColon": (304, 3, 1, 1, 1),
        "TabIndex": (307, 261, 3, 4, 4),
        "Picture": (308, 0, 3, 4, 4),
        "LayoutCachedLeft": (314, 587, 3, 4, 4),
        "LayoutCachedTop": (315, 588, 3, 4, 4),
        "LayoutCachedWidth": (316, 589, 3, 4, 4),
        "LayoutCachedHeight": (317, 590, 3, 4, 4),
        "ThemeFontIndex": (325, 616, 4, 0, 4),
        "BackThemeColorIndex": (326, 617, 4, 0, 4),
        "BorderThemeColorIndex": (329, 620, 4, 0, 4),
        "BorderShade": (331, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (334, 626, 4, 0, 4),
        "GridlineShade": (336, 628, 6, 0, 4),
    },
    "BoundObjectFrame": {
        "AddColon": (5, 3, 1, 1, 0),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "BorderLineStyle": (56, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "LabelX": (103, 52, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "BorderColor": (156, 8, 4, 4, 4),
        "BackColor": (159, 28, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "Tag": (239, 266, 12, 4, 0),
        "ControlTipText": (244, 317, 10, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (291, 0, 3, 4, 4),
        "LayoutCachedLeft": (294, 587, 3, 4, 4),
        "LayoutCachedTop": (295, 588, 3, 4, 4),
        "LayoutCachedWidth": (296, 589, 3, 4, 4),
        "LayoutCachedHeight": (297, 590, 3, 4, 4),
        "BackThemeColorIndex": (303, 617, 4, 0, 4),
        "BorderThemeColorIndex": (306, 620, 4, 0, 4),
        "BorderShade": (308, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (311, 626, 4, 0, 4),
        "GridlineShade": (313, 628, 6, 0, 4),
    },
    "Chart": {
        "OldBorderStyle": (48, 329, 2, 1, 1),
        "BorderLineStyle": (49, 11, 2, 1, 1),
        "Top": (96, 141, 3, 4, 2),
        "Left": (97, 54, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "Name": (221, 20, 10, 4, 0),
        "GridlineColor": (260, 463, 4, 4, 4),
        "OverlapFlags": (274, 159, 2, 1, 1),
        "GUID": (276, 376, 9, 0, 16),
        "GridlineThemeColorIndex": (287, 626, 4, 0, 4),
        "GridlineShade": (289, 628, 6, 0, 4),
        "BackThemeColorIndex": (295, 617, 4, 0, 4),
        "TabIndex": (307, 261, 3, 4, 4),
        "ThemeFontIndex": (379, 616, 4, 0, 4),
    },
    "CheckBox": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "OverlapFlags": (50, 159, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "LabelX": (103, 52, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "BorderColor": (158, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "ValidationRule": (223, 145, 12, 4, 0),
        "ValidationText": (224, 61, 10, 4, 0),
        "DefaultValue": (230, 23, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (290, 0, 3, 4, 4),
        "LayoutCachedLeft": (293, 587, 3, 4, 4),
        "LayoutCachedTop": (294, 588, 3, 4, 4),
        "LayoutCachedWidth": (295, 589, 3, 4, 4),
        "LayoutCachedHeight": (296, 590, 3, 4, 4),
        "BorderThemeColorIndex": (302, 620, 4, 0, 4),
        "BorderShade": (304, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (307, 626, 4, 0, 4),
        "GridlineShade": (309, 628, 6, 0, 4),
    },
    "ComboBox": {
        "AddColon": (8, 3, 1, 1, 0),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "TextFontFamily": (58, 243, 2, 1, 1),
        "BorderLineStyle": (59, 11, 2, 1, 1),
        "DecimalPlaces": (66, 71, 2, 1, 1),
        "TextAlign": (68, 379, 2, 1, 1),
        "ListRows": (97, 153, 3, 4, 2),
        "ListWidth": (98, 154, 3, 4, 2),
        "Left": (99, 54, 3, 4, 2),
        "Top": (100, 141, 3, 4, 2),
        "Width": (101, 150, 3, 4, 2),
        "Height": (102, 44, 3, 4, 2),
        "LabelX": (106, 52, 3, 4, 2),
        "FontSize": (108, 35, 3, 4, 2),
        "FontWeight": (109, 37, 3, 4, 2),
        "TabIndex": (110, 261, 3, 4, 2),
        "BackColor": (157, 28, 4, 4, 4),
        "BorderColor": (158, 8, 4, 4, 4),
        "ForeColor": (160, 204, 4, 4, 4),
        "GUID": (190, 376, 9, 0, 16),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "RowSourceType": (222, 93, 10, 4, 0),
        "RowSource": (223, 91, 12, 4, 0),
        "ColumnWidths": (224, 18, 10, 4, 0),
        "StatusBarText": (225, 135, 10, 4, 0),
        "ValidationRule": (226, 145, 12, 4, 0),
        "ValidationText": (227, 61, 10, 4, 0),
        "DefaultValue": (233, 23, 12, 4, 0),
        "FontName": (234, 34, 10, 64, 0),
        "Tag": (245, 266, 12, 4, 0),
        "ControlTipText": (249, 317, 10, 4, 0),
        "Format": (250, 38, 10, 4, 0),
        "InputMask": (251, 72, 10, 4, 0),
        "GridlineColor": (272, 463, 4, 4, 4),
        "LeftMargin": (302, 384, 3, 4, 2),
        "TopMargin": (303, 385, 3, 4, 2),
        "RightMargin": (304, 388, 3, 4, 2),
        "BottomMargin": (305, 389, 3, 4, 2),
        "Picture": (306, 0, 3, 4, 4),
        "LayoutCachedLeft": (310, 587, 3, 4, 4),
        "LayoutCachedTop": (311, 588, 3, 4, 4),
        "LayoutCachedWidth": (312, 589, 3, 4, 4),
        "LayoutCachedHeight": (313, 590, 3, 4, 4),
        "ThemeFontIndex": (323, 616, 4, 0, 4),
        "BackThemeColorIndex": (324, 617, 4, 0, 4),
        "BorderThemeColorIndex": (327, 620, 4, 0, 4),
        "BorderShade": (329, 622, 6, 0, 4),
        "ForeThemeColorIndex": (330, 623, 4, 0, 4),
        "GridlineThemeColorIndex": (334, 626, 4, 0, 4),
        "GridlineShade": (336, 628, 6, 0, 4),
    },
    "CommandButton": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (49, 159, 2, 1, 1),
        "TextFontFamily": (53, 243, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "FontSize": (103, 35, 3, 4, 2),
        "FontWeight": (104, 37, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "ForeColor": (157, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Caption": (221, 17, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "FontName": (228, 34, 10, 64, 0),
        "Tag": (238, 266, 12, 4, 0),
        "ControlTipText": (241, 317, 10, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (289, 0, 3, 4, 4),
        "LayoutCachedLeft": (293, 587, 3, 4, 4),
        "LayoutCachedTop": (294, 588, 3, 4, 4),
        "LayoutCachedWidth": (295, 589, 3, 4, 4),
        "LayoutCachedHeight": (296, 590, 3, 4, 4),
        "ForeThemeColorIndex": (304, 623, 4, 0, 4),
        "ForeTint": (305, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (311, 626, 4, 0, 4),
        "GridlineShade": (313, 628, 6, 0, 4),
        "Gradient": (318, 693, 4, 0, 4),
        "BackColor": (319, 28, 4, 4, 4),
        "BackThemeColorIndex": (320, 617, 4, 0, 4),
        "OldBorderStyle": (323, 329, 2, 1, 1),
        "BorderLineStyle": (324, 11, 2, 1, 1),
        "BorderWidth": (325, 10, 2, 1, 1),
        "BorderColor": (326, 8, 4, 4, 4),
        "BorderThemeColorIndex": (327, 620, 4, 0, 4),
        "BorderTint": (328, 621, 6, 0, 4),
        "ThemeFontIndex": (330, 616, 4, 0, 4),
        "HoverColor": (331, 653, 4, 4, 4),
        "PressedColor": (335, 657, 4, 4, 4),
        "BottomPadding": (356, 701, 4, 0, 4),
        "TopPadding": (357, 700, 4, 0, 4),
        "LeftPadding": (358, 702, 4, 0, 4),
        "RightPadding": (359, 703, 4, 0, 4),
    },
    "CustomControl": {
        "OldBorderStyle": (52, 329, 2, 1, 1),
        "OverlapFlags": (54, 159, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (101, 261, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (240, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (284, 0, 3, 4, 4),
        "LayoutCachedLeft": (287, 587, 3, 4, 4),
        "LayoutCachedTop": (288, 588, 3, 4, 4),
        "LayoutCachedWidth": (289, 589, 3, 4, 4),
        "LayoutCachedHeight": (290, 590, 3, 4, 4),
        "BackThemeColorIndex": (296, 617, 4, 0, 4),
        "BorderThemeColorIndex": (299, 620, 4, 0, 4),
        "BorderShade": (301, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (304, 626, 4, 0, 4),
        "GridlineShade": (306, 628, 6, 0, 4),
    },
    "Detail": {
        "Height": (96, 44, 3, 4, 2),
        "Name": (223, 20, 10, 4, 0),
        "GUID": (231, 376, 9, 0, 16),
        "AutoHeight": (256, 476, 1, 1, 1),
        "AlternateBackColor": (267, 572, 4, 4, 4),
        "BackThemeColorIndex": (271, 617, 4, 0, 4),
    },
    "EdgeBrowser": {
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (100, 261, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (233, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (284, 0, 3, 4, 4),
        "LayoutCachedLeft": (287, 587, 3, 4, 4),
        "LayoutCachedTop": (288, 588, 3, 4, 4),
        "LayoutCachedWidth": (289, 589, 3, 4, 4),
        "LayoutCachedHeight": (290, 590, 3, 4, 4),
        "BackThemeColorIndex": (307, 617, 4, 0, 4),
        "BorderThemeColorIndex": (310, 620, 4, 0, 4),
        "BorderShade": (312, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (313, 626, 4, 0, 4),
        "GridlineShade": (315, 628, 6, 0, 4),
    },
    "Image": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "BorderLineStyle": (53, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Tag": (228, 266, 12, 4, 0),
        "ControlTipText": (231, 317, 10, 4, 0),
        "GUID": (235, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "ControlSource": (280, 27, 12, 4, 0),
        "Picture": (282, 0, 3, 4, 4),
        "LayoutCachedLeft": (285, 587, 3, 4, 4),
        "LayoutCachedTop": (286, 588, 3, 4, 4),
        "LayoutCachedWidth": (287, 589, 3, 4, 4),
        "LayoutCachedHeight": (288, 590, 3, 4, 4),
        "TabIndex": (289, 261, 3, 4, 4),
        "BackThemeColorIndex": (295, 617, 4, 0, 4),
        "BorderThemeColorIndex": (298, 620, 4, 0, 4),
        "BorderShade": (300, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (305, 626, 4, 0, 4),
        "GridlineShade": (307, 628, 6, 0, 4),
    },
    "Label": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "TextFontFamily": (56, 243, 2, 1, 1),
        "BorderLineStyle": (57, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "FontSize": (100, 35, 3, 4, 2),
        "FontWeight": (101, 37, 3, 4, 2),
        "LeftMargin": (103, 384, 3, 4, 2),
        "TopMargin": (104, 385, 3, 4, 2),
        "LineSpacing": (105, 386, 3, 4, 2),
        "RightMargin": (106, 388, 3, 4, 2),
        "BottomMargin": (107, 389, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "ForeColor": (158, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Caption": (221, 17, 12, 4, 0),
        "FontName": (222, 34, 10, 64, 0),
        "Tag": (228, 266, 12, 4, 0),
        "ControlTipText": (231, 317, 10, 4, 0),
        "GUID": (234, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "LayoutCachedLeft": (282, 587, 3, 4, 4),
        "LayoutCachedTop": (283, 588, 3, 4, 4),
        "LayoutCachedWidth": (284, 589, 3, 4, 4),
        "LayoutCachedHeight": (285, 590, 3, 4, 4),
        "ThemeFontIndex": (291, 616, 4, 0, 4),
        "BackThemeColorIndex": (292, 617, 4, 0, 4),
        "BorderThemeColorIndex": (295, 620, 4, 0, 4),
        "BorderTint": (296, 621, 6, 0, 4),
        "ForeThemeColorIndex": (298, 623, 4, 0, 4),
        "ForeTint": (299, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (305, 626, 4, 0, 4),
        "GridlineShade": (307, 628, 6, 0, 4),
    },
    "Line": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OldBorderStyle": (49, 329, 2, 1, 1),
        "BorderWidth": (50, 10, 2, 1, 1),
        "OverlapFlags": (51, 159, 2, 1, 1),
        "BorderLineStyle": (52, 11, 2, 1, 1),
        "SpecialEffect": (53, 4, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BorderColor": (156, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Tag": (221, 266, 12, 4, 0),
        "GUID": (223, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "LayoutCachedLeft": (276, 587, 3, 4, 4),
        "LayoutCachedTop": (277, 588, 3, 4, 4),
        "LayoutCachedWidth": (278, 589, 3, 4, 4),
        "LayoutCachedHeight": (279, 590, 3, 4, 4),
        "BorderThemeColorIndex": (285, 620, 4, 0, 4),
        "GridlineThemeColorIndex": (290, 626, 4, 0, 4),
        "GridlineShade": (292, 628, 6, 0, 4),
    },
    "ListBox": {
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OverlapFlags": (51, 159, 2, 1, 1),
        "TextFontFamily": (55, 243, 2, 1, 1),
        "OldBorderStyle": (61, 329, 2, 1, 1),
        "BorderWidth": (62, 10, 2, 1, 1),
        "BorderLineStyle": (63, 11, 2, 1, 1),
        "TextAlign": (65, 379, 2, 1, 1),
        "Left": (97, 54, 3, 4, 2),
        "Top": (98, 141, 3, 4, 2),
        "Width": (99, 150, 3, 4, 2),
        "Height": (100, 44, 3, 4, 2),
        "LabelX": (104, 52, 3, 4, 2),
        "FontSize": (106, 35, 3, 4, 2),
        "FontWeight": (107, 37, 3, 4, 2),
        "TabIndex": (108, 261, 3, 4, 2),
        "BackColor": (157, 28, 4, 4, 4),
        "ForeColor": (159, 204, 4, 4, 4),
        "BorderColor": (160, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "RowSourceType": (222, 93, 10, 4, 0),
        "RowSource": (223, 91, 12, 4, 0),
        "ColumnWidths": (224, 18, 10, 4, 0),
        "StatusBarText": (225, 135, 10, 4, 0),
        "ValidationRule": (226, 145, 12, 4, 0),
        "ValidationText": (227, 61, 10, 4, 0),
        "DefaultValue": (233, 23, 12, 4, 0),
        "FontName": (234, 34, 10, 64, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ControlTipText": (247, 317, 10, 4, 0),
        "GUID": (248, 376, 9, 0, 16),
        "GridlineColor": (270, 463, 4, 4, 4),
        "Picture": (296, 0, 3, 4, 4),
        "LayoutCachedLeft": (300, 587, 3, 4, 4),
        "LayoutCachedTop": (301, 588, 3, 4, 4),
        "LayoutCachedWidth": (302, 589, 3, 4, 4),
        "LayoutCachedHeight": (303, 590, 3, 4, 4),
        "ThemeFontIndex": (312, 616, 4, 0, 4),
        "BackThemeColorIndex": (313, 617, 4, 0, 4),
        "BorderThemeColorIndex": (316, 620, 4, 0, 4),
        "BorderShade": (318, 622, 6, 0, 4),
        "ForeThemeColorIndex": (319, 623, 4, 0, 4),
        "ForeTint": (320, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (324, 626, 4, 0, 4),
        "GridlineShade": (326, 628, 6, 0, 4),
    },
    "NavigationControl": {
        "OldBorderStyle": (50, 329, 2, 1, 1),
        "BorderWidth": (51, 10, 2, 1, 1),
        "OverlapFlags": (52, 159, 2, 1, 1),
        "BorderLineStyle": (57, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "ForeColor": (159, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (239, 376, 9, 0, 16),
        "Picture": (287, 0, 3, 4, 4),
        "LayoutCachedLeft": (290, 587, 3, 4, 4),
        "LayoutCachedTop": (291, 588, 3, 4, 4),
        "LayoutCachedWidth": (292, 589, 3, 4, 4),
        "LayoutCachedHeight": (293, 590, 3, 4, 4),
        "BackThemeColorIndex": (299, 617, 4, 0, 4),
        "BorderThemeColorIndex": (302, 620, 4, 0, 4),
        "GridlineColor": (336, 463, 4, 4, 4),
        "GridlineThemeColorIndex": (337, 626, 4, 0, 4),
        "GridlineShade": (339, 628, 6, 0, 4),
    },
    "ObjectFrame": {
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (51, 4, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "OverlapFlags": (55, 159, 2, 1, 1),
        "BorderLineStyle": (56, 11, 2, 1, 1),
        "Left": (97, 54, 3, 4, 2),
        "Top": (98, 141, 3, 4, 2),
        "Width": (99, 150, 3, 4, 2),
        "Height": (100, 44, 3, 4, 2),
        "TabIndex": (103, 261, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "ForeColor": (159, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Tag": (239, 266, 12, 4, 0),
        "ControlTipText": (244, 317, 10, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "GridlineColor": (269, 463, 4, 4, 4),
        "Picture": (285, 0, 3, 4, 4),
        "LayoutCachedLeft": (288, 587, 3, 4, 4),
        "LayoutCachedTop": (289, 588, 3, 4, 4),
        "LayoutCachedWidth": (290, 589, 3, 4, 4),
        "LayoutCachedHeight": (291, 590, 3, 4, 4),
        "ThemeFontIndex": (297, 616, 4, 0, 4),
        "BackThemeColorIndex": (298, 617, 4, 0, 4),
        "BorderThemeColorIndex": (301, 620, 4, 0, 4),
        "BorderShade": (303, 622, 6, 0, 4),
        "ForeThemeColorIndex": (304, 623, 4, 0, 4),
        "GridlineThemeColorIndex": (309, 626, 4, 0, 4),
        "GridlineShade": (311, 628, 6, 0, 4),
    },
    "OptionButton": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "OverlapFlags": (50, 159, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "LabelX": (103, 52, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "BorderColor": (158, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "ValidationRule": (223, 145, 12, 4, 0),
        "ValidationText": (224, 61, 10, 4, 0),
        "DefaultValue": (230, 23, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (290, 0, 3, 4, 4),
        "LayoutCachedLeft": (293, 587, 3, 4, 4),
        "LayoutCachedTop": (294, 588, 3, 4, 4),
        "LayoutCachedWidth": (295, 589, 3, 4, 4),
        "LayoutCachedHeight": (296, 590, 3, 4, 4),
        "BorderThemeColorIndex": (302, 620, 4, 0, 4),
        "BorderShade": (304, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (307, 626, 4, 0, 4),
        "GridlineShade": (309, 628, 6, 0, 4),
    },
    "OptionGroup": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "BorderLineStyle": (56, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "ValidationRule": (223, 145, 12, 4, 0),
        "ValidationText": (224, 61, 10, 4, 0),
        "DefaultValue": (230, 23, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (286, 0, 3, 4, 4),
        "LayoutCachedLeft": (289, 587, 3, 4, 4),
        "LayoutCachedTop": (290, 588, 3, 4, 4),
        "LayoutCachedWidth": (291, 589, 3, 4, 4),
        "LayoutCachedHeight": (292, 590, 3, 4, 4),
        "BackThemeColorIndex": (298, 617, 4, 0, 4),
        "BorderThemeColorIndex": (301, 620, 4, 0, 4),
        "BorderShade": (303, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (306, 626, 4, 0, 4),
        "GridlineShade": (308, 628, 6, 0, 4),
    },
    "Page": {
        "OverlapFlags": (49, 159, 2, 1, 1),
        "Caption": (232, 17, 12, 4, 0),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BorderColor": (156, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (234, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "LayoutCachedLeft": (283, 587, 3, 4, 4),
        "LayoutCachedTop": (284, 588, 3, 4, 4),
        "LayoutCachedWidth": (285, 589, 3, 4, 4),
        "LayoutCachedHeight": (286, 590, 3, 4, 4),
        "BorderThemeColorIndex": (292, 620, 4, 0, 4),
        "BorderShade": (294, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (297, 626, 4, 0, 4),
        "GridlineShade": (299, 628, 6, 0, 4),
        "BottomPadding": (305, 701, 4, 0, 4),
        "TopPadding": (306, 700, 4, 0, 4),
        "LeftPadding": (307, 702, 4, 0, 4),
        "RightPadding": (308, 703, 4, 0, 4),
    },
    "PageBreak": {
        "OverlapFlags": (48, 159, 2, 1, 1),
        "Top": (97, 141, 3, 4, 2),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (223, 376, 9, 0, 16),
    },
    "Rectangle": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "BorderLineStyle": (54, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Tag": (226, 266, 12, 4, 0),
        "GUID": (228, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "LayoutCachedLeft": (281, 587, 3, 4, 4),
        "LayoutCachedTop": (282, 588, 3, 4, 4),
        "LayoutCachedWidth": (283, 589, 3, 4, 4),
        "LayoutCachedHeight": (284, 590, 3, 4, 4),
        "BackThemeColorIndex": (290, 617, 4, 0, 4),
        "BorderThemeColorIndex": (293, 620, 4, 0, 4),
        "BorderShade": (295, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (298, 626, 4, 0, 4),
        "GridlineShade": (300, 628, 6, 0, 4),
    },
    "Subform": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (51, 159, 2, 1, 1),
        "OldBorderStyle": (52, 329, 2, 1, 1),
        "SpecialEffect": (53, 4, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (100, 261, 3, 4, 2),
        "BorderColor": (156, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "StatusBarText": (224, 135, 10, 4, 0),
        "Tag": (227, 266, 12, 4, 0),
        "GUID": (229, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (276, 0, 3, 4, 4),
        "LayoutCachedLeft": (279, 587, 3, 4, 4),
        "LayoutCachedTop": (280, 588, 3, 4, 4),
        "LayoutCachedWidth": (281, 589, 3, 4, 4),
        "LayoutCachedHeight": (282, 590, 3, 4, 4),
        "BorderThemeColorIndex": (288, 620, 4, 0, 4),
        "GridlineThemeColorIndex": (290, 626, 4, 0, 4),
        "GridlineShade": (292, 628, 6, 0, 4),
        "BorderShade": (293, 622, 6, 0, 4),
    },
    "Tab": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (49, 159, 2, 1, 1),
        "TextFontFamily": (52, 243, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "FontSize": (100, 35, 3, 4, 2),
        "FontWeight": (101, 37, 3, 4, 2),
        "TabIndex": (102, 261, 3, 4, 2),
        "Name": (220, 20, 10, 4, 0),
        "StatusBarText": (221, 135, 10, 4, 0),
        "FontName": (223, 34, 10, 64, 0),
        "Tag": (232, 266, 12, 4, 0),
        "GUID": (235, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "Picture": (282, 0, 3, 4, 4),
        "LayoutCachedLeft": (286, 587, 3, 4, 4),
        "LayoutCachedTop": (287, 588, 3, 4, 4),
        "LayoutCachedWidth": (288, 589, 3, 4, 4),
        "LayoutCachedHeight": (289, 590, 3, 4, 4),
        "ThemeFontIndex": (295, 616, 4, 0, 4),
        "GridlineThemeColorIndex": (298, 626, 4, 0, 4),
        "GridlineShade": (300, 628, 6, 0, 4),
        "BackColor": (306, 28, 4, 4, 4),
        "BackThemeColorIndex": (307, 617, 4, 0, 4),
        "OldBorderStyle": (310, 329, 2, 1, 1),
        "BorderLineStyle": (311, 11, 2, 1, 1),
        "BorderColor": (313, 8, 4, 4, 4),
        "BorderThemeColorIndex": (314, 620, 4, 0, 4),
        "BorderTint": (315, 621, 6, 0, 4),
        "HoverColor": (317, 653, 4, 4, 4),
        "PressedColor": (321, 657, 4, 4, 4),
        "ForeColor": (333, 204, 4, 4, 4),
        "ForeThemeColorIndex": (334, 623, 4, 0, 4),
        "ForeTint": (335, 624, 6, 0, 4),
    },
    "TextBox": {
        "AddColon": (8, 3, 1, 1, 0),
        "DecimalPlaces": (48, 71, 2, 1, 1),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (52, 4, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "OverlapFlags": (55, 159, 2, 1, 1),
        "TextFontFamily": (60, 243, 2, 1, 1),
        "BorderLineStyle": (61, 11, 2, 1, 1),
        "TextAlign": (70, 379, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "LabelX": (103, 52, 3, 4, 2),
        "FontSize": (105, 35, 3, 4, 2),
        "FontWeight": (106, 37, 3, 4, 2),
        "TabIndex": (107, 261, 3, 4, 2),
        "LeftMargin": (109, 384, 3, 4, 2),
        "TopMargin": (110, 385, 3, 4, 2),
        "LineSpacing": (111, 386, 3, 4, 2),
        "RightMargin": (112, 388, 3, 4, 2),
        "BottomMargin": (113, 389, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "ForeColor": (159, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "Format": (222, 38, 10, 4, 0),
        "StatusBarText": (223, 135, 10, 4, 0),
        "ValidationRule": (224, 145, 12, 4, 0),
        "ValidationText": (225, 61, 10, 4, 0),
        "DefaultValue": (231, 23, 12, 4, 0),
        "FontName": (232, 34, 10, 64, 0),
        "InputMask": (233, 72, 10, 4, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ControlTipText": (248, 317, 10, 4, 0),
        "GUID": (250, 376, 9, 0, 16),
        "GridlineColor": (272, 463, 4, 4, 4),
        "Picture": (299, 0, 3, 4, 4),
        "LayoutCachedLeft": (302, 587, 3, 4, 4),
        "LayoutCachedTop": (303, 588, 3, 4, 4),
        "LayoutCachedWidth": (304, 589, 3, 4, 4),
        "LayoutCachedHeight": (305, 590, 3, 4, 4),
        "BackThemeColorIndex": (313, 617, 4, 0, 4),
        "BorderThemeColorIndex": (316, 620, 4, 0, 4),
        "BorderShade": (318, 622, 6, 0, 4),
        "ThemeFontIndex": (319, 616, 4, 0, 4),
        "ForeThemeColorIndex": (320, 623, 4, 0, 4),
        "ForeTint": (321, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (327, 626, 4, 0, 4),
        "GridlineShade": (329, 628, 6, 0, 4),
    },
    "ToggleButton": {
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (49, 159, 2, 1, 1),
        "TextFontFamily": (53, 243, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "FontSize": (105, 35, 3, 4, 2),
        "TabIndex": (107, 261, 3, 4, 2),
        "ForeColor": (158, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "ValidationRule": (223, 145, 12, 4, 0),
        "ValidationText": (224, 61, 10, 4, 0),
        "DefaultValue": (230, 23, 12, 4, 0),
        "Caption": (231, 17, 12, 4, 0),
        "FontName": (233, 34, 10, 64, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ControlTipText": (247, 317, 10, 4, 0),
        "GUID": (249, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (291, 0, 3, 4, 4),
        "LayoutCachedLeft": (294, 587, 3, 4, 4),
        "LayoutCachedTop": (295, 588, 3, 4, 4),
        "LayoutCachedWidth": (296, 589, 3, 4, 4),
        "LayoutCachedHeight": (297, 590, 3, 4, 4),
        "ForeThemeColorIndex": (303, 623, 4, 0, 4),
        "ForeTint": (304, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (308, 626, 4, 0, 4),
        "GridlineShade": (310, 628, 6, 0, 4),
        "BackColor": (316, 28, 4, 4, 4),
        "BackThemeColorIndex": (317, 617, 4, 0, 4),
        "OldBorderStyle": (320, 329, 2, 1, 1),
        "BorderLineStyle": (321, 11, 2, 1, 1),
        "BorderWidth": (322, 10, 2, 1, 1),
        "BorderColor": (323, 8, 4, 4, 4),
        "BorderThemeColorIndex": (324, 620, 4, 0, 4),
        "BorderTint": (325, 621, 6, 0, 4),
        "ThemeFontIndex": (327, 616, 4, 0, 4),
        "HoverColor": (328, 653, 4, 4, 4),
        "PressedColor": (332, 657, 4, 4, 4),
        "BottomPadding": (353, 701, 4, 0, 4),
        "TopPadding": (354, 700, 4, 0, 4),
        "LeftPadding": (355, 702, 4, 0, 4),
        "RightPadding": (356, 703, 4, 0, 4),
    },
    "WebBrowser": {
        "OldBorderStyle": (52, 329, 2, 1, 1),
        "OverlapFlags": (54, 159, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (101, 261, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (239, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "Picture": (289, 0, 3, 4, 4),
        "LayoutCachedLeft": (292, 587, 3, 4, 4),
        "LayoutCachedTop": (293, 588, 3, 4, 4),
        "LayoutCachedWidth": (294, 589, 3, 4, 4),
        "LayoutCachedHeight": (295, 590, 3, 4, 4),
        "BackThemeColorIndex": (315, 617, 4, 0, 4),
        "BorderThemeColorIndex": (318, 620, 4, 0, 4),
        "BorderShade": (320, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (321, 626, 4, 0, 4),
        "GridlineShade": (323, 628, 6, 0, 4),
    },
    "_Design": {
        "Width": (99, 150, 3, 4, 2),
        "Left": (104, 54, 3, 4, 2),
        "Top": (105, 141, 3, 4, 2),
        "GUID": (208, 376, 9, 0, 16),
        "Caption": (221, 17, 12, 4, 0),
        "BorderThemeColorIndex": (407, 620, 4, 0, 4),
        "ThemeFontIndex": (410, 616, 4, 0, 4),
        "ForeThemeColorIndex": (411, 623, 4, 0, 4),
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
    "CustomControl",
    "Attachment",
    "WebBrowser",
)

#: Read but not written.  Each carries records this project cannot name,
#: and a navigation control, chart or Edge browser repeats one code at two
#: ids, which a table keyed by property name cannot express -- and one
#: written without the data source that gives it its content would not be
#: a working control anyway.
READ_ONLY_TYPES = ("NavigationControl", "Chart", "EdgeBrowser")
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
        if control_type in READ_ONLY_TYPES:
            raise AccessError(
                f"a {control_type} is read but not written: it carries records "
                f"this project cannot name, and one written without its data "
                f"source would not work anyway"
            )
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



def encode_property(slot: tuple[int, int, int, int, int], value: object) -> bytes:
    """One property's value as its slot says to write it."""
    _ident, _code, value_type, _width, size = slot
    if value_type in TEXT_VALUE_TYPES:
        if not isinstance(value, str):
            raise AccessError(f"this property takes text, not {type(value).__name__}")
        return value.encode("utf-16-le")
    if value_type == 11:
        if not isinstance(value, (bytes, bytearray)):
            raise AccessError("this property takes raw bytes")
        return bytes(value)
    if value_type == 9:
        if not isinstance(value, (bytes, bytearray)) or len(value) != GUID_LENGTH:
            raise AccessError(f"a GUID is {GUID_LENGTH} bytes")
        return bytes(value)
    if value_type == 6:
        return struct.pack("<f", float(value))  # pyright: ignore[reportArgumentType]
    if value_type == 8:
        return struct.pack("<d", float(value))  # pyright: ignore[reportArgumentType]
    if isinstance(value, bool):
        number = 1 if value else 0
    elif isinstance(value, int):
        number = value
    else:
        raise AccessError(f"this property takes a number, not {type(value).__name__}")
    width = size or 1
    try:
        return number.to_bytes(width, "little", signed=number < 0)
    except OverflowError as exc:
        raise AccessError(f"{number} does not fit the property's {width} bytes") from exc


def set_property(blob: bytes, target: str | None, name: str, value: object) -> bytes:
    """A design with one property of one of its objects changed.

    `target` names a control or a section; `None` means the design itself.
    A property the object already carries keeps its record's id and length
    conventions; one it does not gets the id its own type's schema gives
    it, so the records stay in the order Access reads them.
    """
    header, objects, trailer = parse_design(blob)
    def wanted(obj: DesignObject) -> bool:
        return obj.name == target if target is not None else obj.type is None

    at = next((i for i, obj in enumerate(objects) if wanted(obj)), None)
    if at is None:
        raise AccessError(
            f"this design has no object named {target!r}" if target is not None
            else "this design has no object of its own"
        )
    obj = objects[at]
    kind = DESIGN_OBJECT if obj.type is None else CONTROL_TYPES.get(obj.type)
    slots = PROPERTY_SLOTS.get(kind or "")
    if slots is None:
        raise AccessError(f"no property slots were measured for a {kind}")
    slot = slots.get(name)
    if slot is None:
        known = ", ".join(sorted(slots))
        raise AccessError(f"a {kind} has no {name!r} to set; it has: {known}")
    ident, code, value_type, width, _size = slot
    raw = encode_property(slot, value)

    kept = [r for r in obj.records if r.code != code]
    existing = next((r for r in obj.records if r.code == code), None)
    if existing is not None:
        # Keep the record where it is: its id is the schema's answer for
        # this object, whatever the table says.
        replaced = DesignRecord(existing.id, code, existing.value_type, existing.width, raw)
    else:
        replaced = DesignRecord(ident, code, value_type, width, raw)
    records = tuple(sorted((*kept, replaced), key=lambda r: r.id))
    rebuilt = (*objects[:at], DesignObject(obj.marker, obj.type, obj.code, records), *objects[at + 1 :])
    return build_design(header, rebuilt, trailer)


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
