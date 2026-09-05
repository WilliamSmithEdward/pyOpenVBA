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
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import NamedTuple

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
    127: "EmptyCell",
    128: "WebBrowser",
    129: "NavigationControl",
    130: "NavigationButton",
    133: "Chart",
    134: "EdgeBrowser",
    152: "Detail",
    153: "HeaderSection",
    154: "FooterSection",
    155: "PageHeaderSection",
    156: "PageFooterSection",
    157: "GroupHeaderSection",
    158: "GroupFooterSection",
}
#: The types above that are sections: a form's header and footer, a
#: report's header, footer and group levels, and the page sections and
#: detail both have.  A section is named by its own record (FormHeader
#: on a form, ReportHeader on a report, GroupHeader0 and up).
SECTION_TYPES = frozenset(range(152, 159))
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
        return self.type in SECTION_TYPES

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
    "Unidentified450": 450,
    "Unidentified596": 596,
    "Unidentified597": 597,
    "Unidentified600": 600,
    "Unidentified640": 640,
    "Unidentified779": 779,
    "Unidentified783": 783,
    "Unidentified787": 787,
    "Unidentified792": 792,
    "Unidentified799": 799,
    "Unidentified806": 806,
    "Transparent": 46,
    "TextFontCharSet": 193,
    "HelpContextId": 219,
    "PictureType": 341,
    "HyperlinkAddress": 343,
    "HyperlinkSubAddress": 347,
    "LeftPadding": 455,
    "TopPadding": 456,
    "RightPadding": 457,
    "BottomPadding": 458,
    "GridlineStyleLeft": 459,
    "GridlineStyleTop": 460,
    "GridlineStyleRight": 461,
    "GridlineStyleBottom": 462,
    "GridlineWidthLeft": 464,
    "GridlineWidthTop": 465,
    "GridlineWidthRight": 466,
    "GridlineWidthBottom": 467,
    "CursorOnHover": 585,
    "PictureCaptionArrangement": 594,
    "Alignment": 595,
    "ScrollTop": 604,
    "ScrollLeft": 605,
    "OnDocumentComplete": 609,
    "OnNavigateError": 611,
    "BackTint": 618,
    "BackShade": 619,
    "ForeShade": 625,
    "Span": 646,
    "UseTheme": 649,
    "Shape": 650,
    "Bevel": 651,
    "HoverShade": 656,
    "HoverForeShade": 664,
    "PressedForeShade": 668,
    "SoftEdges": 670,
    "QuickStyleMask": 691,
    "ChartType": 754,
    "HasLegend": 755,
    "HasTitle": 756,
    "ChartTitle": 757,
    "HasAxisTitles": 758,
    "CategoryAxisTitle": 759,
    "PrimaryValuesAxisTitle": 760,
    "LegendPosition": 769,
    "ChartTitleFontSize": 796,
    "HasSubtitle": 803,
    "ChartSubtitle": 804,
    "AutoRepeat": 5,
    "BoundColumn": 13,
    "CanGrow": 14,
    "CanShrink": 16,
    "Enabled": 25,
    "FontItalic": 33,
    "FontUnderline": 36,
    "LimitToList": 67,
    "AutoExpand": 68,
    "ColumnHeads": 131,
    "TextAlign": 136,
    "Visible": 148,
    "ScrollBars": 152,
    "Default": 220,
    "Cancel": 221,
    "TabStop": 262,
    "MultiSelect": 293,
    "TripleState": 294,
    "FilterLookup": 305,
    "PictureAlignment": 309,
    "PictureTiling": 313,
    "AllowAutoCorrect": 340,
    "IMEMode": 356,
    "ReadingOrder": 367,
    "KeyboardLanguage": 368,
    "ScrollBarAlign": 371,
    "NumeralShapes": 372,
    "IsHyperlink": 396,
    "TextFormat": 477,
    "VerticalAnchor": 478,
    "HorizontalAnchor": 479,
    "ShowDatePicker": 487,
    "DisplayAsHyperlink": 592,
    "LinkChildFields": 49,
    "LinkMasterFields": 50,
    "Locked": 56,
    "OnClick": 126,
    "ShortcutMenuBar": 304,
    "IMESentenceMode": 379,
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
    # Named by differencing on a command button and a toggle button
    # (docs/research/access_designs): the property that moved alone, or
    # the one left once the colour it resets was set aside.
    "Glow": 652,
    "HoverThemeColorIndex": 654,
    "HoverTint": 655,
    "PressedThemeColorIndex": 658,
    "PressedTint": 659,
    "PressedShade": 660,
    "HoverForeColor": 661,
    "HoverForeThemeColorIndex": 662,
    "HoverForeTint": 663,
    "PressedForeColor": 665,
    "PressedForeThemeColorIndex": 666,
    "PressedForeTint": 667,
    "Shadow": 669,
    "QuickStyle": 690,
    "ControlSource": 27,
    "RowSourceType": 93,
    "TabIndex": 261,
    "Picture": 0,
    # Not properties: a button or page without Glow or Shadow carries
    # 2, 2, 1, 1 here, and setting either drops the four and writes the
    # paddings instead (measured by differencing), so these are the
    # insets Access keeps for an effect that is off.
    "Unidentified700": 700,
    "Unidentified701": 701,
    "Unidentified702": 702,
    "Unidentified703": 703,
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
        "IMESentenceMode": (70, 379, 2, 1),
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
        "Unidentified701": (356, 701, 4, 0),
        "Unidentified700": (357, 700, 4, 0),
        "Unidentified702": (358, 702, 4, 0),
        "Unidentified703": (359, 703, 4, 0),
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
        "Unidentified701": (353, 701, 4, 0),
        "Unidentified700": (354, 700, 4, 0),
        "Unidentified702": (355, 702, 4, 0),
        "Unidentified703": (356, 703, 4, 0),
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
        "IMESentenceMode": (65, 379, 2, 1),
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
        "IMESentenceMode": (68, 379, 2, 1),
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
        "Unidentified701": (305, 701, 4, 0),
        "Unidentified700": (306, 700, 4, 0),
        "Unidentified702": (307, 702, 4, 0),
        "Unidentified703": (308, 703, 4, 0),
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
    # A chart carries no picture and no cached layout, and its records
    # run past 750 into the axis and title colours.
    "Chart": {
        "Top": (96, 141, 3, 4),
        "Left": (97, 54, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "Name": (221, 20, 10, 4),
        "TopPadding": (265, 456, 3, 4),
        "BottomPadding": (266, 458, 3, 4),
        "HorizontalAnchor": (269, 479, 3, 4),
        "OverlapFlags": (274, 159, 2, 1),
        "GUID": (276, 376, 9, 0),
        "Unidentified600": (279, 600, 3, 4),
        "Unidentified450": (280, 450, 3, 4),
        "Unidentified596": (283, 596, 3, 4),
        "Unidentified597": (284, 597, 3, 4),
        "TabIndex": (307, 261, 3, 4),
        "Unidentified779": (360, 779, 4, 4),
        "Unidentified783": (364, 783, 4, 4),
        "Unidentified787": (368, 787, 4, 4),
        "Unidentified792": (373, 792, 4, 4),
        "ChartTitleFontSize": (377, 796, 3, 2),
        "Unidentified799": (381, 799, 4, 4),
        "Unidentified806": (388, 806, 4, 4),
    },
    # Code 450 appears twice at two ids, which is why the slots are keyed
    # by code where no name is established rather than by name alone.
    "EdgeBrowser": {
        "OverlapFlags": (53, 159, 2, 1),
        "Left": (96, 54, 3, 4),
        "Top": (97, 141, 3, 4),
        "Width": (98, 150, 3, 4),
        "Height": (99, 44, 3, 4),
        "TabIndex": (100, 261, 3, 4),
        "Name": (220, 20, 10, 4),
        "GUID": (233, 376, 9, 0),
        "Unidentified450": (259, 450, 3, 4),
        "TopPadding": (261, 456, 3, 4),
        "BottomPadding": (263, 458, 3, 4),
        "HorizontalAnchor": (273, 479, 3, 4),
        "Picture": (284, 0, 3, 4),
        "LayoutCachedLeft": (287, 587, 3, 4),
        "LayoutCachedTop": (288, 588, 3, 4),
        "LayoutCachedWidth": (289, 589, 3, 4),
        "LayoutCachedHeight": (290, 590, 3, 4),
        "Unidentified600": (301, 600, 3, 4),
        "Unidentified596": (302, 596, 3, 4),
        "Unidentified597": (303, 597, 3, 4),
        "Unidentified450_2": (322, 450, 3, 4),
    },
}

#: The design's own object, which has no control type of its own.
DESIGN_OBJECT = "_Design"
#: Value types whose length comes from the value rather than the slot.
VARIABLE_VALUE_TYPES = (10, 11, 12)
#: Both string types, which differ in what Access lets you put in them --
#: 12 takes an expression -- and not in how the bytes are written.
TEXT_VALUE_TYPES = (10, 12)

def property_code(name: str) -> int | None:
    """The code a slot key names.

    A control can carry one code at two ids -- an Edge browser carries 450
    twice -- so the second slot's key ends in `_2`.  Strip that and the
    key names the same property.
    """
    if name in PROPERTY_CODES:
        return PROPERTY_CODES[name]
    base = name.rsplit("_", 1)[0]
    return PROPERTY_CODES.get(base) if base != name else None


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
        "Visible": (0, 148, 1, 1, 0),
        "PictureTiling": (1, 313, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "BorderLineStyle": (53, 11, 2, 1, 1),
        "PictureAlignment": (54, 309, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ControlSource": (221, 27, 12, 4, 0),
        "OnClick": (231, 126, 12, 4, 0),
        "Tag": (239, 266, 12, 4, 0),
        "ShortcutMenuBar": (241, 304, 10, 4, 0),
        "ControlTipText": (242, 317, 10, 4, 0),
        "GUID": (243, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "HorizontalAnchor": (292, 479, 3, 4, 4),
        "VerticalAnchor": (293, 478, 3, 4, 4),
        "Enabled": (296, 25, 1, 1, 1),
        "Locked": (297, 56, 1, 1, 1),
        "LabelX": (302, 52, 3, 4, 4),
        "AddColon": (304, 3, 1, 1, 1),
        "TabStop": (306, 262, 1, 1, 1),
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
        "StatusBarText": (339, 135, 10, 4, 0),
    },
    "BoundObjectFrame": {
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "AddColon": (5, 3, 1, 1, 0),
        "TabStop": (7, 262, 1, 1, 0),
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
        "OnClick": (234, 126, 12, 4, 0),
        "Tag": (239, 266, 12, 4, 0),
        "ShortcutMenuBar": (243, 304, 10, 4, 0),
        "ControlTipText": (244, 317, 10, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "TopPadding": (265, 456, 3, 4, 4),
        "BottomPadding": (266, 458, 3, 4, 4),
        "HorizontalAnchor": (269, 479, 3, 4, 4),
        "RowSource": (271, 91, 12, 4, 0),
        "OverlapFlags": (274, 159, 2, 1, 1),
        "GUID": (276, 376, 9, 0, 16),
        "GridlineThemeColorIndex": (287, 626, 4, 0, 4),
        "GridlineShade": (289, 628, 6, 0, 4),
        "ChartType": (293, 754, 3, 4, 4),
        "BackThemeColorIndex": (295, 617, 4, 0, 4),
        "HasLegend": (298, 755, 1, 1, 1),
        "HasTitle": (299, 756, 1, 1, 1),
        "ChartTitle": (300, 757, 10, 4, 0),
        "HasAxisTitles": (301, 758, 1, 1, 1),
        "CategoryAxisTitle": (302, 759, 10, 4, 0),
        "PrimaryValuesAxisTitle": (303, 760, 10, 4, 0),
        "LinkChildFields": (304, 49, 10, 4, 0),
        "TabIndex": (307, 261, 3, 4, 4),
        "LegendPosition": (346, 769, 3, 4, 4),
        "ChartTitleFontSize": (377, 796, 3, 2, 2),
        "ThemeFontIndex": (379, 616, 4, 0, 4),
        "HasSubtitle": (385, 803, 1, 1, 1),
        "ChartSubtitle": (386, 804, 10, 4, 0),
    },
    "CheckBox": {
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "TabStop": (7, 262, 1, 1, 0),
        "TripleState": (8, 294, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "OverlapFlags": (50, 159, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "ReadingOrder": (56, 367, 2, 1, 1),
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
        "OnClick": (239, 126, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ShortcutMenuBar": (242, 304, 10, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "ColumnHeads": (0, 131, 1, 1, 0),
        "LimitToList": (1, 67, 1, 1, 0),
        "Visible": (2, 148, 1, 1, 0),
        "Enabled": (3, 25, 1, 1, 0),
        "Locked": (4, 56, 1, 1, 0),
        "AddColon": (8, 3, 1, 1, 0),
        "FontItalic": (9, 33, 1, 1, 0),
        "FontUnderline": (10, 36, 1, 1, 0),
        "AutoExpand": (11, 68, 1, 1, 0),
        "TabStop": (12, 262, 1, 1, 0),
        "AllowAutoCorrect": (14, 340, 1, 1, 0),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "TextAlign": (57, 136, 2, 1, 1),
        "TextFontFamily": (58, 243, 2, 1, 1),
        "BorderLineStyle": (59, 11, 2, 1, 1),
        "IMEMode": (60, 356, 2, 1, 1),
        "ReadingOrder": (61, 367, 2, 1, 1),
        "KeyboardLanguage": (62, 368, 2, 1, 1),
        "ScrollBarAlign": (64, 371, 2, 1, 1),
        "DecimalPlaces": (66, 71, 2, 1, 1),
        "NumeralShapes": (67, 372, 2, 1, 1),
        "IMESentenceMode": (68, 379, 2, 1, 1),
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
        "BoundColumn": (156, 13, 4, 4, 4),
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
        "OnClick": (243, 126, 12, 4, 0),
        "Tag": (245, 266, 12, 4, 0),
        "ShortcutMenuBar": (248, 304, 10, 4, 0),
        "ControlTipText": (249, 317, 10, 4, 0),
        "Format": (250, 38, 10, 4, 0),
        "InputMask": (251, 72, 10, 4, 0),
        "GridlineColor": (272, 463, 4, 4, 4),
        "CanGrow": (277, 14, 1, 1, 1),
        "CanShrink": (278, 16, 1, 1, 1),
        "HorizontalAnchor": (280, 479, 3, 4, 4),
        "VerticalAnchor": (281, 478, 3, 4, 4),
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
        "ForeShade": (332, 625, 6, 0, 4),
        "GridlineThemeColorIndex": (334, 626, 4, 0, 4),
        "GridlineShade": (336, 628, 6, 0, 4),
    },
    "CommandButton": {
        "Transparent": (0, 46, 1, 1, 0),
        "AutoRepeat": (1, 5, 1, 1, 0),
        "Default": (2, 220, 1, 1, 0),
        "Cancel": (3, 221, 1, 1, 0),
        "Visible": (4, 148, 1, 1, 0),
        "Enabled": (5, 25, 1, 1, 0),
        "FontItalic": (8, 33, 1, 1, 0),
        "FontUnderline": (9, 36, 1, 1, 0),
        "TabStop": (10, 262, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (49, 159, 2, 1, 1),
        "TextFontCharSet": (52, 193, 2, 1, 1),
        "TextFontFamily": (53, 243, 2, 1, 1),
        "ReadingOrder": (54, 367, 2, 1, 1),
        "PictureType": (55, 341, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "FontSize": (103, 35, 3, 4, 2),
        "FontWeight": (104, 37, 3, 4, 2),
        "TabIndex": (105, 261, 3, 4, 2),
        "HelpContextId": (156, 219, 4, 4, 4),
        "ForeColor": (157, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Caption": (221, 17, 12, 4, 0),
        "StatusBarText": (222, 135, 10, 4, 0),
        "OnClick": (223, 126, 12, 4, 0),
        "FontName": (228, 34, 10, 64, 0),
        "Tag": (238, 266, 12, 4, 0),
        "ShortcutMenuBar": (240, 304, 10, 4, 0),
        "ControlTipText": (241, 317, 10, 4, 0),
        "HyperlinkAddress": (243, 343, 12, 4, 0),
        "HyperlinkSubAddress": (244, 347, 12, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "LeftPadding": (260, 455, 3, 4, 4),
        "TopPadding": (261, 456, 3, 4, 4),
        "RightPadding": (262, 457, 3, 4, 4),
        "BottomPadding": (263, 458, 3, 4, 4),
        "GridlineStyleTop": (265, 460, 2, 1, 1),
        "GridlineColor": (268, 463, 4, 4, 4),
        "GridlineWidthTop": (270, 465, 2, 1, 1),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
        "Picture": (289, 0, 3, 4, 4),
        "CursorOnHover": (292, 585, 3, 2, 2),
        "LayoutCachedLeft": (293, 587, 3, 4, 4),
        "LayoutCachedTop": (294, 588, 3, 4, 4),
        "LayoutCachedWidth": (295, 589, 3, 4, 4),
        "LayoutCachedHeight": (296, 590, 3, 4, 4),
        "PictureCaptionArrangement": (297, 594, 2, 1, 1),
        "Alignment": (298, 595, 2, 1, 1),
        "ForeThemeColorIndex": (304, 623, 4, 0, 4),
        "ForeTint": (305, 624, 6, 0, 4),
        "ForeShade": (306, 625, 6, 0, 4),
        "GridlineThemeColorIndex": (311, 626, 4, 0, 4),
        "GridlineShade": (313, 628, 6, 0, 4),
        "UseTheme": (314, 649, 1, 0, 1),
        "Shape": (315, 650, 4, 0, 4),
        "Bevel": (316, 651, 4, 0, 4),
        "Gradient": (318, 693, 4, 0, 4),
        "BackColor": (319, 28, 4, 4, 4),
        "BackThemeColorIndex": (320, 617, 4, 0, 4),
        "BackTint": (321, 618, 6, 0, 4),
        "BackShade": (322, 619, 6, 0, 4),
        "OldBorderStyle": (323, 329, 2, 1, 1),
        "BorderLineStyle": (324, 11, 2, 1, 1),
        "BorderWidth": (325, 10, 2, 1, 1),
        "BorderColor": (326, 8, 4, 4, 4),
        "BorderThemeColorIndex": (327, 620, 4, 0, 4),
        "BorderTint": (328, 621, 6, 0, 4),
        "ThemeFontIndex": (330, 616, 4, 0, 4),
        "HoverColor": (331, 653, 4, 4, 4),
        "HoverThemeColorIndex": (332, 654, 4, 0, 4),
        "HoverShade": (334, 656, 6, 0, 4),
        "PressedColor": (335, 657, 4, 4, 4),
        "PressedThemeColorIndex": (336, 658, 4, 0, 4),
        "HoverForeColor": (339, 661, 4, 4, 4),
        "HoverForeThemeColorIndex": (340, 662, 4, 0, 4),
        "PressedForeColor": (343, 665, 4, 4, 4),
        "PressedForeThemeColorIndex": (344, 666, 4, 0, 4),
        "HoverForeShade": (342, 664, 6, 0, 4),
        "PressedForeShade": (346, 668, 6, 0, 4),
        "SoftEdges": (349, 670, 4, 0, 4),
        "QuickStyleMask": (351, 691, 4, 0, 4),
        "Unidentified701": (356, 701, 4, 0, 4),
        "Unidentified700": (357, 700, 4, 0, 4),
        "Unidentified702": (358, 702, 4, 0, 4),
        "Unidentified703": (359, 703, 4, 0, 4),
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
        "CanGrow": (2, 14, 1, 1, 0),
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
        "TopPadding": (261, 456, 3, 4, 4),
        "BottomPadding": (263, 458, 3, 4, 4),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "Picture": (284, 0, 3, 4, 4),
        "LayoutCachedLeft": (287, 587, 3, 4, 4),
        "LayoutCachedTop": (288, 588, 3, 4, 4),
        "LayoutCachedWidth": (289, 589, 3, 4, 4),
        "LayoutCachedHeight": (290, 590, 3, 4, 4),
        "ScrollTop": (293, 604, 4, 4, 4),
        "ScrollLeft": (294, 605, 4, 4, 4),
        "OnDocumentComplete": (297, 609, 12, 4, 0),
        "OnNavigateError": (299, 611, 12, 4, 0),
        "BackThemeColorIndex": (307, 617, 4, 0, 4),
        "BorderThemeColorIndex": (310, 620, 4, 0, 4),
        "BorderShade": (312, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (313, 626, 4, 0, 4),
        "GridlineShade": (315, 628, 6, 0, 4),
    },
    "Image": {
        "Visible": (0, 148, 1, 1, 0),
        "PictureTiling": (1, 313, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "BorderLineStyle": (53, 11, 2, 1, 1),
        "PictureAlignment": (55, 309, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "OnClick": (221, 126, 12, 4, 0),
        "Tag": (228, 266, 12, 4, 0),
        "ShortcutMenuBar": (230, 304, 10, 4, 0),
        "ControlTipText": (231, 317, 10, 4, 0),
        "GUID": (235, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "HorizontalAnchor": (272, 479, 3, 4, 4),
        "VerticalAnchor": (273, 478, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "FontItalic": (1, 33, 1, 1, 0),
        "FontUnderline": (2, 36, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "BackStyle": (50, 29, 2, 1, 1),
        "OldBorderStyle": (51, 329, 2, 1, 1),
        "BorderWidth": (52, 10, 2, 1, 1),
        "OverlapFlags": (53, 159, 2, 1, 1),
        "TextAlign": (55, 136, 2, 1, 1),
        "TextFontFamily": (56, 243, 2, 1, 1),
        "BorderLineStyle": (57, 11, 2, 1, 1),
        "ReadingOrder": (58, 367, 2, 1, 1),
        "NumeralShapes": (59, 372, 2, 1, 1),
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
        "OnClick": (223, 126, 12, 4, 0),
        "Tag": (228, 266, 12, 4, 0),
        "ShortcutMenuBar": (230, 304, 10, 4, 0),
        "ControlTipText": (231, 317, 10, 4, 0),
        "GUID": (234, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
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
        "HorizontalAnchor": (272, 479, 3, 4, 4),
        "VerticalAnchor": (273, 478, 3, 4, 4),
        "LayoutCachedLeft": (276, 587, 3, 4, 4),
        "LayoutCachedTop": (277, 588, 3, 4, 4),
        "LayoutCachedWidth": (278, 589, 3, 4, 4),
        "LayoutCachedHeight": (279, 590, 3, 4, 4),
        "BorderThemeColorIndex": (285, 620, 4, 0, 4),
        "GridlineThemeColorIndex": (290, 626, 4, 0, 4),
        "GridlineShade": (292, 628, 6, 0, 4),
    },
    "ListBox": {
        "ColumnHeads": (0, 131, 1, 1, 0),
        "Visible": (1, 148, 1, 1, 0),
        "Enabled": (2, 25, 1, 1, 0),
        "Locked": (3, 56, 1, 1, 0),
        "FontItalic": (8, 33, 1, 1, 0),
        "FontUnderline": (9, 36, 1, 1, 0),
        "TabStop": (10, 262, 1, 1, 0),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OverlapFlags": (51, 159, 2, 1, 1),
        "TextFontFamily": (55, 243, 2, 1, 1),
        "MultiSelect": (56, 293, 2, 1, 1),
        "IMEMode": (57, 356, 2, 1, 1),
        "ReadingOrder": (58, 367, 2, 1, 1),
        "ScrollBarAlign": (59, 371, 2, 1, 1),
        "OldBorderStyle": (61, 329, 2, 1, 1),
        "BorderWidth": (62, 10, 2, 1, 1),
        "BorderLineStyle": (63, 11, 2, 1, 1),
        "NumeralShapes": (64, 372, 2, 1, 1),
        "IMESentenceMode": (65, 379, 2, 1, 1),
        "Left": (97, 54, 3, 4, 2),
        "Top": (98, 141, 3, 4, 2),
        "Width": (99, 150, 3, 4, 2),
        "Height": (100, 44, 3, 4, 2),
        "LabelX": (104, 52, 3, 4, 2),
        "FontSize": (106, 35, 3, 4, 2),
        "FontWeight": (107, 37, 3, 4, 2),
        "TabIndex": (108, 261, 3, 4, 2),
        "BoundColumn": (156, 13, 4, 4, 4),
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
        "OnClick": (243, 126, 12, 4, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ShortcutMenuBar": (246, 304, 10, 4, 0),
        "ControlTipText": (247, 317, 10, 4, 0),
        "GUID": (248, 376, 9, 0, 16),
        "GridlineColor": (270, 463, 4, 4, 4),
        "HorizontalAnchor": (275, 479, 3, 4, 4),
        "VerticalAnchor": (276, 478, 3, 4, 4),
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
        "BackColor": (156, 28, 4, 4, 4),
        "BorderColor": (157, 8, 4, 4, 4),
        "HelpContextId": (158, 219, 4, 4, 4),
        "ForeColor": (159, 204, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "ShortcutMenuBar": (237, 304, 10, 4, 0),
        "GUID": (239, 376, 9, 0, 16),
        "LeftPadding": (261, 455, 3, 4, 4),
        "TopPadding": (262, 456, 3, 4, 4),
        "RightPadding": (263, 457, 3, 4, 4),
        "BottomPadding": (264, 458, 3, 4, 4),
        "GridlineStyleLeft": (265, 459, 2, 1, 1),
        "GridlineStyleTop": (266, 460, 2, 1, 1),
        "GridlineStyleRight": (267, 461, 2, 1, 1),
        "GridlineStyleBottom": (268, 462, 2, 1, 1),
        "GridlineWidthLeft": (269, 464, 2, 1, 1),
        "GridlineWidthTop": (270, 465, 2, 1, 1),
        "GridlineWidthRight": (271, 466, 2, 1, 1),
        "GridlineWidthBottom": (272, 467, 2, 1, 1),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "Picture": (287, 0, 3, 4, 4),
        "LayoutCachedLeft": (290, 587, 3, 4, 4),
        "LayoutCachedTop": (291, 588, 3, 4, 4),
        "LayoutCachedWidth": (292, 589, 3, 4, 4),
        "LayoutCachedHeight": (293, 590, 3, 4, 4),
        "BackThemeColorIndex": (299, 617, 4, 0, 4),
        "BackTint": (300, 618, 6, 0, 4),
        "BackShade": (301, 619, 6, 0, 4),
        "BorderThemeColorIndex": (302, 620, 4, 0, 4),
        "BorderTint": (303, 621, 6, 0, 4),
        "BorderShade": (304, 622, 6, 0, 4),
        "Span": (314, 646, 3, 4, 4),
        "GridlineColor": (336, 463, 4, 4, 4),
        "GridlineThemeColorIndex": (337, 626, 4, 0, 4),
        "GridlineShade": (339, 628, 6, 0, 4),
    },
    "ObjectFrame": {
        "Visible": (0, 148, 1, 1, 0),
        "TabStop": (5, 262, 1, 1, 0),
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
        "OnClick": (232, 126, 12, 4, 0),
        "Tag": (239, 266, 12, 4, 0),
        "ShortcutMenuBar": (243, 304, 10, 4, 0),
        "ControlTipText": (244, 317, 10, 4, 0),
        "GUID": (245, 376, 9, 0, 16),
        "GridlineColor": (269, 463, 4, 4, 4),
        "HorizontalAnchor": (274, 479, 3, 4, 4),
        "VerticalAnchor": (275, 478, 3, 4, 4),
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
        "ForeShade": (306, 625, 6, 0, 4),
        "GridlineThemeColorIndex": (309, 626, 4, 0, 4),
        "GridlineShade": (311, 628, 6, 0, 4),
    },
    "OptionButton": {
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "TabStop": (7, 262, 1, 1, 0),
        "TripleState": (8, 294, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "SpecialEffect": (49, 4, 2, 1, 1),
        "OverlapFlags": (50, 159, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "ReadingOrder": (56, 367, 2, 1, 1),
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
        "OnClick": (239, 126, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ShortcutMenuBar": (242, 304, 10, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "TabStop": (7, 262, 1, 1, 0),
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
        "OnClick": (239, 126, 12, 4, 0),
        "Tag": (240, 266, 12, 4, 0),
        "ShortcutMenuBar": (242, 304, 10, 4, 0),
        "ControlTipText": (243, 317, 10, 4, 0),
        "GUID": (244, 376, 9, 0, 16),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "BorderColor": (156, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Caption": (232, 17, 12, 4, 0),
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
        "Unidentified701": (305, 701, 4, 0, 4),
        "Unidentified700": (306, 700, 4, 0, 4),
        "Unidentified702": (307, 702, 4, 0, 4),
        "Unidentified703": (308, 703, 4, 0, 4),
    },
    "PageBreak": {
        "OverlapFlags": (48, 159, 2, 1, 1),
        "Top": (97, 141, 3, 4, 2),
        "Name": (220, 20, 10, 4, 0),
        "GUID": (223, 376, 9, 0, 16),
    },
    "Rectangle": {
        "Visible": (0, 148, 1, 1, 0),
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
        "OnClick": (221, 126, 12, 4, 0),
        "Tag": (226, 266, 12, 4, 0),
        "GUID": (228, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "HorizontalAnchor": (272, 479, 3, 4, 4),
        "VerticalAnchor": (273, 478, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "CanShrink": (4, 16, 1, 1, 0),
        "TabStop": (7, 262, 1, 1, 0),
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
        "LinkChildFields": (222, 49, 10, 4, 0),
        "LinkMasterFields": (223, 50, 10, 4, 0),
        "StatusBarText": (224, 135, 10, 4, 0),
        "Tag": (227, 266, 12, 4, 0),
        "GUID": (229, 376, 9, 0, 16),
        "LeftPadding": (260, 455, 3, 4, 4),
        "TopPadding": (261, 456, 3, 4, 4),
        "RightPadding": (262, 457, 3, 4, 4),
        "BottomPadding": (263, 458, 3, 4, 4),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "FontItalic": (2, 33, 1, 1, 0),
        "FontUnderline": (3, 36, 1, 1, 0),
        "TabStop": (4, 262, 1, 1, 0),
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
        "OnClick": (230, 126, 12, 4, 0),
        "Tag": (232, 266, 12, 4, 0),
        "ShortcutMenuBar": (234, 304, 10, 4, 0),
        "GUID": (235, 376, 9, 0, 16),
        "GridlineColor": (267, 463, 4, 4, 4),
        "HorizontalAnchor": (272, 479, 3, 4, 4),
        "VerticalAnchor": (273, 478, 3, 4, 4),
        "Picture": (282, 0, 3, 4, 4),
        "LayoutCachedLeft": (286, 587, 3, 4, 4),
        "LayoutCachedTop": (287, 588, 3, 4, 4),
        "LayoutCachedWidth": (288, 589, 3, 4, 4),
        "LayoutCachedHeight": (289, 590, 3, 4, 4),
        "ThemeFontIndex": (295, 616, 4, 0, 4),
        "GridlineThemeColorIndex": (298, 626, 4, 0, 4),
        "GridlineShade": (300, 628, 6, 0, 4),
        "UseTheme": (301, 649, 1, 0, 1),
        "Shape": (302, 650, 4, 0, 4),
        "BackColor": (306, 28, 4, 4, 4),
        "BackThemeColorIndex": (307, 617, 4, 0, 4),
        "BackShade": (309, 619, 6, 0, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "CanGrow": (4, 14, 1, 1, 0),
        "CanShrink": (5, 16, 1, 1, 0),
        "AddColon": (8, 3, 1, 1, 0),
        "FontItalic": (9, 33, 1, 1, 0),
        "FontUnderline": (10, 36, 1, 1, 0),
        "TabStop": (12, 262, 1, 1, 0),
        "AllowAutoCorrect": (16, 340, 1, 1, 0),
        "IsHyperlink": (18, 396, 1, 1, 0),
        "DecimalPlaces": (48, 71, 2, 1, 1),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "ScrollBars": (50, 152, 2, 1, 1),
        "SpecialEffect": (52, 4, 2, 1, 1),
        "OldBorderStyle": (53, 329, 2, 1, 1),
        "BorderWidth": (54, 10, 2, 1, 1),
        "OverlapFlags": (55, 159, 2, 1, 1),
        "TextAlign": (59, 136, 2, 1, 1),
        "TextFontFamily": (60, 243, 2, 1, 1),
        "BorderLineStyle": (61, 11, 2, 1, 1),
        "IMEMode": (62, 356, 2, 1, 1),
        "ReadingOrder": (63, 367, 2, 1, 1),
        "KeyboardLanguage": (64, 368, 2, 1, 1),
        "ScrollBarAlign": (66, 371, 2, 1, 1),
        "BackStyle": (67, 29, 2, 1, 1),
        "FilterLookup": (68, 305, 2, 1, 1),
        "NumeralShapes": (69, 372, 2, 1, 1),
        "IMESentenceMode": (70, 379, 2, 1, 1),
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
        "OnClick": (242, 126, 12, 4, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ShortcutMenuBar": (247, 304, 10, 4, 0),
        "ControlTipText": (248, 317, 10, 4, 0),
        "GUID": (250, 376, 9, 0, 16),
        "TopPadding": (265, 456, 3, 4, 4),
        "BottomPadding": (267, 458, 3, 4, 4),
        "GridlineColor": (272, 463, 4, 4, 4),
        "TextFormat": (277, 477, 2, 1, 1),
        "HorizontalAnchor": (278, 479, 3, 4, 4),
        "VerticalAnchor": (279, 478, 3, 4, 4),
        "ShowDatePicker": (280, 487, 3, 4, 4),
        "Picture": (299, 0, 3, 4, 4),
        "LayoutCachedLeft": (302, 587, 3, 4, 4),
        "LayoutCachedTop": (303, 588, 3, 4, 4),
        "LayoutCachedWidth": (304, 589, 3, 4, 4),
        "LayoutCachedHeight": (305, 590, 3, 4, 4),
        "DisplayAsHyperlink": (307, 592, 3, 4, 4),
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
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "Locked": (2, 56, 1, 1, 0),
        "FontItalic": (7, 33, 1, 1, 0),
        "FontUnderline": (8, 36, 1, 1, 0),
        "TabStop": (9, 262, 1, 1, 0),
        "TripleState": (10, 294, 1, 1, 0),
        "DisplayWhen": (48, 149, 2, 1, 1),
        "OverlapFlags": (49, 159, 2, 1, 1),
        "TextFontFamily": (53, 243, 2, 1, 1),
        "ReadingOrder": (54, 367, 2, 1, 1),
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
        "OnClick": (242, 126, 12, 4, 0),
        "Tag": (244, 266, 12, 4, 0),
        "ShortcutMenuBar": (246, 304, 10, 4, 0),
        "ControlTipText": (247, 317, 10, 4, 0),
        "GUID": (249, 376, 9, 0, 16),
        "LeftPadding": (260, 455, 3, 4, 4),
        "TopPadding": (261, 456, 3, 4, 4),
        "RightPadding": (262, 457, 3, 4, 4),
        "BottomPadding": (263, 458, 3, 4, 4),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
        "Picture": (291, 0, 3, 4, 4),
        "LayoutCachedLeft": (294, 587, 3, 4, 4),
        "LayoutCachedTop": (295, 588, 3, 4, 4),
        "LayoutCachedWidth": (296, 589, 3, 4, 4),
        "LayoutCachedHeight": (297, 590, 3, 4, 4),
        "ForeThemeColorIndex": (303, 623, 4, 0, 4),
        "ForeTint": (304, 624, 6, 0, 4),
        "GridlineThemeColorIndex": (308, 626, 4, 0, 4),
        "GridlineShade": (310, 628, 6, 0, 4),
        "UseTheme": (311, 649, 1, 0, 1),
        "Shape": (312, 650, 4, 0, 4),
        "Bevel": (313, 651, 4, 0, 4),
        "BackColor": (316, 28, 4, 4, 4),
        "BackThemeColorIndex": (317, 617, 4, 0, 4),
        "BackTint": (318, 618, 6, 0, 4),
        "OldBorderStyle": (320, 329, 2, 1, 1),
        "BorderLineStyle": (321, 11, 2, 1, 1),
        "BorderWidth": (322, 10, 2, 1, 1),
        "BorderColor": (323, 8, 4, 4, 4),
        "BorderThemeColorIndex": (324, 620, 4, 0, 4),
        "BorderTint": (325, 621, 6, 0, 4),
        "ThemeFontIndex": (327, 616, 4, 0, 4),
        "HoverColor": (328, 653, 4, 4, 4),
        "PressedColor": (332, 657, 4, 4, 4),
        "Unidentified701": (353, 701, 4, 0, 4),
        "Unidentified700": (354, 700, 4, 0, 4),
        "Unidentified702": (355, 702, 4, 0, 4),
        "Unidentified703": (356, 703, 4, 0, 4),
    },
    "WebBrowser": {
        "Visible": (0, 148, 1, 1, 0),
        "Enabled": (1, 25, 1, 1, 0),
        "TabStop": (3, 262, 1, 1, 0),
        "DisplayWhen": (49, 149, 2, 1, 1),
        "SpecialEffect": (50, 4, 2, 1, 1),
        "OldBorderStyle": (52, 329, 2, 1, 1),
        "BorderWidth": (53, 10, 2, 1, 1),
        "OverlapFlags": (54, 159, 2, 1, 1),
        "BorderLineStyle": (55, 11, 2, 1, 1),
        "Left": (96, 54, 3, 4, 2),
        "Top": (97, 141, 3, 4, 2),
        "Width": (98, 150, 3, 4, 2),
        "Height": (99, 44, 3, 4, 2),
        "TabIndex": (101, 261, 3, 4, 2),
        "BorderColor": (157, 8, 4, 4, 4),
        "Name": (220, 20, 10, 4, 0),
        "Tag": (232, 266, 12, 4, 0),
        "ControlTipText": (237, 317, 10, 4, 0),
        "GUID": (239, 376, 9, 0, 16),
        "StatusBarText": (240, 135, 10, 4, 0),
        "GridlineColor": (268, 463, 4, 4, 4),
        "HorizontalAnchor": (273, 479, 3, 4, 4),
        "VerticalAnchor": (274, 478, 3, 4, 4),
        "Picture": (289, 0, 3, 4, 4),
        "LayoutCachedLeft": (292, 587, 3, 4, 4),
        "LayoutCachedTop": (293, 588, 3, 4, 4),
        "LayoutCachedWidth": (294, 589, 3, 4, 4),
        "LayoutCachedHeight": (295, 590, 3, 4, 4),
        "ControlSource": (296, 27, 12, 4, 0),
        "BackThemeColorIndex": (315, 617, 4, 0, 4),
        "BorderThemeColorIndex": (318, 620, 4, 0, 4),
        "BorderShade": (320, 622, 6, 0, 4),
        "GridlineThemeColorIndex": (321, 626, 4, 0, 4),
        "GridlineShade": (323, 628, 6, 0, 4),
    },
    "_Design": {
        "PictureAlignment": (61, 309, 2, 1, 1),
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
#: Four records Access writes on a new command button, toggle button and
#: page.  They were once taken for the control's padding -- which is codes
#: 455 to 458 -- and what they really are is not established, so they are
#: written exactly as Access writes them and named by their code.
BUTTON_EXTRAS = {
    "CommandButton": {700: 2, 701: 2, 702: 1, 703: 1},
    "ToggleButton": {700: 2, 701: 2, 702: 2, 703: 2},
    "Page": {700: 2, 701: 2, 702: 2, 703: 2},
}
#: What Access writes into a chart's and an Edge browser's own records.
#: Some are not named -- differencing every property those controls expose
#: left them unaccounted for -- and are written anyway, because a control
#: missing records Access gives it is not the control Access made.
#:
#: Three are deliberately left out: 596, 597 and 600 put the control in a
#: layout, which is where Access's designer drops a new one.  A control
#: placed at an explicit position must not claim layout membership --
#: with them written, Access stacks a chart under whatever else claims the
#: same layout instead of leaving it where it was put.
TYPE_EXTRAS: dict[str, dict[str, bytes]] = {
    "Chart": {
        "TopPadding": bytes(4),
        "BottomPadding": bytes(4),
        "Unidentified450": (2).to_bytes(4, "little"),
        "ChartTitleFontSize": (14).to_bytes(2, "little"),
        "Unidentified779": bytes.fromhex("7f7f7f00"),
        "Unidentified783": bytes.fromhex("7f7f7f00"),
        "Unidentified787": bytes.fromhex("7f7f7f00"),
        "Unidentified792": bytes.fromhex("7f7f7f00"),
        "Unidentified799": bytes.fromhex("7f7f7f00"),
        "Unidentified806": bytes.fromhex("7f7f7f00"),
    },
    "EdgeBrowser": {
        "Unidentified450": (2).to_bytes(4, "little"),
        "TopPadding": bytes(4),
        "BottomPadding": bytes(4),
        "Unidentified450_2": (2).to_bytes(4, "little"),
    },
}
#: The overlap flags Access writes, where they are not the usual 85.
OVERLAP_FLAGS = {"Chart": 0x5D, "EdgeBrowser": 0x57, "BoundObjectFrame": 0x5D}
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

#: Read but not written.  A navigation control is not one control: one of
#: its records names a sibling subform by name, and Access builds
#: navigation buttons beside it, so writing the control alone would leave
#: it pointing at a subform that is not there.
READ_ONLY_TYPES = ("NavigationControl", "NavigationButton")
#: What Access sets on a text box, list box or combo box it has just
#: made.  It was long taken for the text alignment, which it is not.
DEFAULT_IME_SENTENCE_MODE = 3
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
                f"a {control_type} is read but not written: it is part of a "
                f"navigation control, which names a sibling subform and the "
                f"buttons beside it, so one written alone would not work"
            )
        raise AccessError(
            f"a {control_type} cannot be written yet; known: {', '.join(sorted(CONTROL_SLOTS))}"
        )
    # The control Access made for the measurement was the first that
    # takes the focus, so it carried no TabIndex; every later one does,
    # at the slot Access's own designs give it (measured: three text
    # boxes written without it came back from Access tabbed 1, 2, 0).
    slots = dict(slots)
    measured = PROPERTY_SLOTS.get(control_type, {})
    if "TabIndex" in measured:
        slots.setdefault("TabIndex", measured["TabIndex"][:4])
    values: dict[str, bytes] = {
        "OverlapFlags": bytes((OVERLAP_FLAGS.get(control_type, DEFAULT_OVERLAP),)),
        "IMESentenceMode": bytes((DEFAULT_IME_SENTENCE_MODE,)),
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
    for code, number in BUTTON_EXTRAS.get(control_type, {}).items():
        values[f"Unidentified{code}"] = number.to_bytes(4, "little")
    values.update(TYPE_EXTRAS.get(control_type, {}))
    present = [(key, value) for key, value in values.items() if key in slots]
    records = tuple(
        _record(slots[key], value)
        for key, value in sorted(present, key=lambda kv: slots[kv[0]][0])
    )
    return DesignObject(None, TYPE_CODES[control_type], None, records)


#: Ahead of its sections a design carries one nameless object per control
#: type it holds: the type's defaults on this design (the property sheet's
#: "Set Control Defaults"), written when the first control of the type is
#: added and never taken away.  Access reads a control's themed properties
#: against them; without the object it ignores a button's UseTheme, colour
#: indexes and gradient and drops them on its next save (measured).  They
#: are constant for a type across every design Access built from the blank
#: template, and differ between a form and a report, so each kind's set is
#: captured from the every-control fixture.
_PROTOTYPES: dict[str, dict[int, tuple[DesignRecord, ...]]] = {}
#: A page break has no defaults object: the every-control form and report
#: each hold one and carry none for it.
DEFAULTS_FREE = frozenset({118})


def prototype_records(kind: str, type_code: int) -> tuple[DesignRecord, ...]:
    """The records of the control-defaults object for `type_code` on a
    design of `kind`; refused for a type the fixture did not hold."""
    if kind not in _PROTOTYPES:
        _header, objects, _trailer = parse_design(template(kind, "prototypes"))
        _PROTOTYPES[kind] = {o.type: o.records for o in objects[1:] if o.type is not None}
    try:
        return _PROTOTYPES[kind][type_code]
    except KeyError:
        name = CONTROL_TYPES.get(type_code, type_code)
        raise AccessError(f"no control defaults were captured for a {name} on a {kind}") from None


def with_prototype(
    objects: tuple[DesignObject, ...], kind: str, type_code: int
) -> tuple[DesignObject, ...]:
    """The design with the control-defaults object for `type_code` in
    place: inserted in type order among the ones there, the run of
    top-level objects (defaults, then sections) re-marked as one group."""
    first_section = next((i for i, o in enumerate(objects) if o.is_section), len(objects))
    ahead = list(objects[1:first_section])
    if type_code in DEFAULTS_FREE or any(o.type == type_code for o in ahead):
        return objects
    fresh = DesignObject(None, type_code, None, prototype_records(kind, type_code))
    at = next((i for i, o in enumerate(ahead) if (o.type or 0) > type_code), len(ahead))
    ahead.insert(at, fresh)
    sections = [(i, o) for i, o in enumerate(objects) if o.is_section]
    top = _placed((*ahead, *(o for _i, o in sections)))
    rest = list(objects[first_section:])
    for position, (index, _o) in enumerate(sections):
        rest[index - first_section] = top[len(ahead) + position]
    return (objects[0], *top[: len(ahead)], *rest)


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


#: Setting one of these in Access writes a second record beside it: the
#: theme index at -1, so the colour or font written is the one used rather
#: than the theme's.  Measured one property at a time on a label, text
#: box, list box, combo box, check box, rectangle, line, toggle button and
#: command button; a control type without the companion's slot (a check box
#: has no fill) takes only the property itself.
COMPANIONS: dict[str, tuple[tuple[str, object], ...]] = {
    "FontName": (("ThemeFontIndex", -1),),
    "ForeColor": (("ForeThemeColorIndex", -1),),
    "BackColor": (("BackThemeColorIndex", -1),),
    "BorderColor": (("BorderThemeColorIndex", -1),),
    "GridlineColor": (("GridlineThemeColorIndex", -1),),
    "HoverColor": (("HoverThemeColorIndex", -1),),
    "PressedColor": (("PressedThemeColorIndex", -1),),
    "HoverForeColor": (("HoverForeThemeColorIndex", -1),),
    "PressedForeColor": (("PressedForeThemeColorIndex", -1),),
}
#: A command button's fill also turns its gradient off (measured; a toggle
#: button's does not).
BUTTON_FILL_COMPANIONS: tuple[tuple[str, object], ...] = (("Gradient", 0),)
#: With a font's name Access writes the font's pitch-and-family byte, read
#: from the installed font; these are the values it wrote for the fonts
#: below (measured on 24 fonts), and a font not listed gets its name alone.
FONT_FAMILIES: dict[str, int] = {
    "Book Antiqua": 18,
    "Cambria": 18,
    "Comic Sans MS": 66,
    "Consolas": 49,
    "Courier New": 49,
    "Garamond": 18,
    "Georgia": 18,
    "Lucida Console": 49,
    "Palatino Linotype": 18,
    "Symbol": 18,
    "Times New Roman": 18,
    "Wingdings": 2,
}
#: For these Access wrote no family byte at all (theirs is the default a
#: variable-pitch sans face takes): Segoe UI, Segoe UI Semibold, Calibri,
#: Arial, Tahoma, Verdana, Trebuchet MS, Century Gothic, Candara, Corbel,
#: Franklin Gothic Medium and Impact.


def _companions(kind: str | None, name: str, value: object) -> tuple[tuple[str, object], ...]:
    out = list(COMPANIONS.get(name, ()))
    if name == "BackColor" and kind == "CommandButton":
        out.extend(BUTTON_FILL_COMPANIONS)
    if name == "FontName" and isinstance(value, str) and value in FONT_FAMILIES:
        out.append(("TextFontFamily", FONT_FAMILIES[value]))
    return tuple(out)


#: Beside the ``RowSourceType`` text, a list or combo box carries the kind
#: as a byte at this slot when it is not the default table or query.
ROW_SOURCE_KIND_CODE = 92
ROW_SOURCE_KIND_SLOT = (48, ROW_SOURCE_KIND_CODE, 2, 1)
ROW_SOURCE_KINDS = {"Value List": 1, "Field List": 10}


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
    if name == "RowSourceType":
        # The text is what the property sheet shows; what Access acts on is
        # a one-byte companion record: 1 for a value list, 10 for a field
        # list, none for a table or query (measured on list and combo
        # boxes Access built with each).
        kept = [r for r in kept if r.code != ROW_SOURCE_KIND_CODE]
        companion = ROW_SOURCE_KINDS.get(str(value))
        if companion is not None:
            kept.append(_record(ROW_SOURCE_KIND_SLOT, bytes((companion,))))
    # A colour or font brings the records Access writes beside it, each at
    # the id the object already uses for it or the schema's id otherwise.
    for extra_name, setting in _companions(kind, name, value):
        extra_slot = slots.get(extra_name)
        if extra_slot is None:
            continue
        kept = [r for r in kept if r.code != extra_slot[1]]
        found = next((r for r in obj.records if r.code == extra_slot[1]), None)
        extra_raw = encode_property(extra_slot, setting)
        kept.append(
            DesignRecord(found.id, extra_slot[1], found.value_type, found.width, extra_raw)
            if found is not None
            else DesignRecord(extra_slot[0], extra_slot[1], extra_slot[2], extra_slot[3], extra_raw)
        )
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
    kind: str = "form",
) -> bytes:
    """A design with one more control on it.

    A control belongs to a section and is written immediately after it,
    before the section that follows.  Adding one rewrites the markers of
    the controls already there, since they depend on how many the section
    holds; see `_placed`.

    `parent` names a control that holds controls of its own -- a tab
    control holding pages -- and the new one joins that group rather than
    the section's.

    The first control of a type also brings the type's control-defaults
    object ahead of the sections, as it does in Access (`with_prototype`).
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
            if o.name and o.type is not None and CONTROL_TYPES.get(o.type) in TABBABLE
        ),
    )
    # The first control of a type brings the type's control-defaults object
    # ahead of the sections, which moves the sections along.
    objects = with_prototype(objects, kind, TYPE_CODES[control_type])
    at = next(i for i, o in enumerate(objects) if o.is_section and o.name == section)
    end = at + 1
    while end < len(objects) and not objects[end].is_section:
        end += 1
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


def _tab_index_of(obj: DesignObject, tab_code: int | None) -> int | None:
    """Where a control sits in the tab order, or `None` for one that never
    takes the focus.  A control with no record sits at 0, which is how
    Access writes the first one."""
    if obj.type is None or CONTROL_TYPES.get(obj.type) not in TABBABLE or tab_code is None:
        return None
    raw = obj.property_value(tab_code)
    return int.from_bytes(raw, "little") if raw else 0


def _closed_up(obj: DesignObject, removed: int, tab_code: int) -> DesignObject:
    """`obj` with its tab index one lower when it followed the removed
    control; a control that lands on 0 loses the record, as Access writes
    the first control without one."""
    raw = obj.property_value(tab_code)
    if obj.type is None or not raw:
        return obj
    index = int.from_bytes(raw, "little")
    if index <= removed:
        return obj
    records = tuple(
        DesignRecord(r.id, r.code, r.value_type, r.width, (index - 1).to_bytes(len(r.value), "little"))
        if r.code == tab_code else r
        for r in obj.records
        if not (r.code == tab_code and index - 1 == 0)
    )
    return DesignObject(obj.marker, obj.type, obj.code, records)


def remove_control(blob: bytes, name: str) -> bytes:
    """A design with the named control taken off, and whatever it holds.

    The controls left at that level are re-marked, since their markers say
    how many there are (see `_placed`), and the tab order closes up behind
    a control that took the focus, as it does when Access deletes one in
    its designer.  A page's own controls are not written here, so a
    control below a page is refused rather than guessed at.
    """
    header, objects, trailer = parse_design(blob)
    tab_code = property_code("TabIndex")
    result: list[DesignObject] | None = None
    removed_tab: int | None = None
    at = 1
    while at < len(objects) and result is None:
        if not objects[at].is_section:
            at += 1
            continue
        end = at + 1
        while end < len(objects) and not objects[end].is_section:
            end += 1
        owners = _top_level(objects, at + 1, end)
        hit = next((pair for pair in owners if objects[pair[0]].name == name), None)
        if hit is not None:
            index, _owned = hit
            removed_tab = _tab_index_of(objects[index], tab_code)
            rest = [pair for pair in owners if pair[0] != index]
            placed = _placed(tuple(objects[i] for i, _ in rest))
            rebuilt: list[DesignObject] = []
            for position, (i, kept) in enumerate(rest):
                rebuilt.append(placed[position])
                rebuilt.extend(objects[i + 1 : i + 1 + kept])
            result = [*objects[: at + 1], *rebuilt, *objects[end:]]
            break
        for index, owned in owners:
            children = objects[index + 1 : index + 1 + owned]
            child = next((k for k, obj in enumerate(children) if obj.name == name), None)
            if child is None:
                continue
            if _children_span(objects, index + 1 + child, index + 1 + owned):
                raise AccessError(f"{name!r} holds controls of its own; take those off first")
            removed_tab = _tab_index_of(children[child], tab_code)
            kept_children = _placed(tuple(obj for k, obj in enumerate(children) if k != child))
            result = [*objects[: index + 1], *kept_children, *objects[index + 1 + owned :]]
            break
        at = end
    if result is None:
        raise AccessError(f"this design has no control named {name!r}")
    if removed_tab is not None and tab_code is not None:
        result = [_closed_up(obj, removed_tab, tab_code) for obj in result]
    return build_design(header, tuple(result), trailer)

#: The ``TypeInfo`` stream beside a design lists its sections and controls
#: as the members VBA sees on the form's class: ``Me.Qty`` compiles, and
#: ``Qty_Click`` binds, only for a name listed here.  After the magic come
#: a kind word, -1, the entry count and the design's CLSID; each entry is
#: then ``<u32 type id><u32 ordinal><name in the code page>00 00``.
#:
#: Access keeps the stream rather than rebuilding it (measured by editing
#: forms in Access and reading the stream back): a new member is appended
#: with the ordinal above the highest present, a removed one drops out
#: while the others keep their ordinals, and a renamed one moves to the
#: end with its ordinal.  Only a copy made with CopyObject is built afresh,
#: controls first in design order and the sections after them.
TYPE_INFO_MAGIC = bytes.fromhex("f7eacdac")
TYPE_INFO_KIND = {"form": 0x96, "report": 0x197}
TYPE_INFO_CLSID_AT = 16
TYPE_INFO_ENTRIES_AT = 32
#: A type id's low byte is the object's type code and its high byte the
#: index Access gives the member's class, read off a form and a report
#: holding one of every section and every control.
TYPE_INFO_SECTIONS = {
    "form": {
        "Detail": 0x1898,
        "HeaderSection": 0x1899,
        "FooterSection": 0x189A,
        "PageHeaderSection": 0x189B,
        "PageFooterSection": 0x189C,
    },
    "report": {
        "Detail": 0x1998,
        "HeaderSection": 0x1999,
        "FooterSection": 0x199A,
        "PageHeaderSection": 0x1F9B,
        "PageFooterSection": 0x1F9C,
        "GroupHeaderSection": 0x199D,
        "GroupFooterSection": 0x199E,
    },
}
#: A form's controls, standing on a section or on a page.
TYPE_INFO_IDS = {
    "Label": 0x0D64,
    "Rectangle": 0x0E65,
    "Line": 0x1666,
    "Image": 0x0F67,
    "CommandButton": 0x0B68,
    "OptionButton": 0x0869,
    "CheckBox": 0x066A,
    "OptionGroup": 0x116B,
    "BoundObjectFrame": 0x026C,
    "TextBox": 0x126D,
    "ListBox": 0x106E,
    "ComboBox": 0x136F,
    "Subform": 0x1470,
    "ObjectFrame": 0x0372,
    "PageBreak": 0x1776,
    "CustomControl": 0x0477,
    "ToggleButton": 0x0A7A,
    "Tab": 0x207B,
    "Page": 0x217C,
    "Attachment": 0x227E,
    "EmptyCell": 0x247F,
    "WebBrowser": 0x2580,
    "NavigationControl": 0x2681,
    "NavigationButton": 0x2782,
    "Chart": 0x2885,
    "EdgeBrowser": 0x2986,
}
#: The same controls held by another control -- a label attached to the
#: control it describes, a button inside an option group -- belong to a
#: class of their own.  A page holds controls without changing them.
TYPE_INFO_HELD_IDS = {
    "Label": 0x0C64,
    "OptionButton": 0x0769,
    "CheckBox": 0x056A,
    "ToggleButton": 0x097A,
}
#: On a report every control measured shares one class index, whatever
#: holds it, except an ActiveX control.
TYPE_INFO_REPORT_IDS = {
    name: 0x1B00 | code
    for code, name in CONTROL_TYPES.items()
    if name
    in (
        "Label",
        "Rectangle",
        "Line",
        "Image",
        "CommandButton",
        "OptionButton",
        "CheckBox",
        "OptionGroup",
        "BoundObjectFrame",
        "TextBox",
        "ListBox",
        "ComboBox",
        "Subform",
        "ObjectFrame",
        "PageBreak",
        "ToggleButton",
        "Tab",
        "Page",
        "Attachment",
        "Chart",
        "EdgeBrowser",
    )
}
#: An ActiveX control's entry carries 36 more bytes on either kind; an
#: empty control, which is all that is written here, carries zeros.
ACTIVEX_CONTROL = "CustomControl"
ACTIVEX_CODE = next(code for code, name in CONTROL_TYPES.items() if name == ACTIVEX_CONTROL)
TYPE_INFO_REPORT_IDS[ACTIVEX_CONTROL] = TYPE_INFO_IDS[ACTIVEX_CONTROL]
TYPE_INFO_ACTIVEX_TAIL = bytes(36)


class TypeInfoEntry(NamedTuple):
    """One member a TypeInfo stream lists."""

    ident: int
    ordinal: int
    name: str
    tail: bytes = b""


class TypeInfoMember(NamedTuple):
    """A member a design has, with the type id Access gives it."""

    name: str
    ident: int
    tail: bytes = b""


def type_info_entries(stream: bytes) -> tuple[TypeInfoEntry, ...]:
    """The members a TypeInfo stream lists, in the order it lists them."""
    if stream[:4] != TYPE_INFO_MAGIC or len(stream) < TYPE_INFO_ENTRIES_AT:
        raise AccessError("this is not a TypeInfo stream")
    (count,) = struct.unpack_from("<I", stream, 12)
    out: list[TypeInfoEntry] = []
    at = TYPE_INFO_ENTRIES_AT
    for _ in range(count):
        if at + 8 > len(stream):
            raise AccessError("this TypeInfo stream ends inside an entry")
        ident, ordinal = struct.unpack_from("<II", stream, at)
        end = stream.find(bytes(2), at + 8)
        if end < 0:
            raise AccessError("this TypeInfo stream ends inside a name")
        name = stream[at + 8 : end].decode("cp1252")
        at = end + 2
        tail = b""
        if ident & 0xFF == ACTIVEX_CODE:
            tail = stream[at : at + len(TYPE_INFO_ACTIVEX_TAIL)]
            at += len(tail)
        out.append(TypeInfoEntry(ident, ordinal, name, tail))
    return tuple(out)


def _holders(objects: tuple[DesignObject, ...]) -> dict[int, int]:
    """What holds each control -- its section, or the control it sits
    inside -- as indexes into `objects`."""
    holders: dict[int, int] = {}

    def walk(start: int, stop: int, holder: int) -> None:
        for index, owned in _top_level(objects, start, stop):
            holders[index] = holder
            walk(index + 1, index + 1 + owned, index)

    at = 1
    while at < len(objects):
        if objects[at].is_section:
            end = at + 1
            while end < len(objects) and not objects[end].is_section:
                end += 1
            walk(at + 1, end, at)
            at = end
        else:
            at += 1
    return holders


def type_info_members(kind: str, objects: tuple[DesignObject, ...]) -> tuple[TypeInfoMember, ...]:
    """The members a design has -- its sections, then its named controls,
    each in design order -- with the type id Access gives each one.  A type
    whose id has not been measured on that kind of design is refused
    rather than guessed."""
    if kind not in TYPE_INFO_KIND:
        raise AccessError(f"kind must be 'form' or 'report', not {kind!r}")
    holders = _holders(objects)
    sections: list[TypeInfoMember] = []
    controls: list[TypeInfoMember] = []
    for index, obj in enumerate(objects):
        name = obj.name
        if index == 0 or not name:
            continue
        type_name = obj.type_name or ""
        if obj.is_section:
            ident = TYPE_INFO_SECTIONS[kind].get(type_name)
        elif kind == "report":
            ident = TYPE_INFO_REPORT_IDS.get(type_name)
        else:
            holder = holders.get(index)
            held = (
                holder is not None
                and not objects[holder].is_section
                and objects[holder].type_name != PAGE
            )
            ident = (TYPE_INFO_HELD_IDS.get(type_name) if held else None) or TYPE_INFO_IDS.get(
                type_name
            )
        if ident is None:
            raise AccessError(
                f"no TypeInfo id has been measured for a {type_name or obj.type} on a {kind}"
            )
        tail = TYPE_INFO_ACTIVEX_TAIL if type_name == ACTIVEX_CONTROL else b""
        (sections if obj.is_section else controls).append(TypeInfoMember(name, ident, tail))
    return (*sections, *controls)


def type_info(
    kind: str,
    objects: tuple[DesignObject, ...],
    existing: bytes,
    *,
    renamed: Mapping[str, str] | None = None,
) -> bytes:
    """The TypeInfo stream after the design changed, carried forward the
    way Access carries it: members the design no longer has drop out, the
    ones `renamed` maps (old name to new) move to the end keeping their
    ordinals, and new members are appended with the ordinals that follow
    the highest present."""
    entries = type_info_entries(existing)
    members = type_info_members(kind, objects)
    wanted = {member.name for member in members}
    renames = dict(renamed or {})
    kept = [e for e in entries if e.name in wanted and e.name not in renames]
    moved = [
        e._replace(name=renames[e.name])
        for e in entries
        if e.name in renames and renames[e.name] in wanted
    ]
    present = {e.name for e in (*kept, *moved)}
    ordinal = max((e.ordinal for e in (*kept, *moved)), default=-1) + 1
    added: list[TypeInfoEntry] = []
    for member in members:
        if member.name not in present:
            added.append(TypeInfoEntry(member.ident, ordinal, member.name, member.tail))
            ordinal += 1
    clsid = existing[TYPE_INFO_CLSID_AT : TYPE_INFO_CLSID_AT + GUID_LENGTH]
    return pack_type_info(kind, clsid, (*kept, *moved, *added))


def pack_type_info(kind: str, clsid: bytes, entries: Sequence[TypeInfoEntry]) -> bytes:
    """A TypeInfo stream listing `entries` in that order."""
    if kind not in TYPE_INFO_KIND:
        raise AccessError(f"kind must be 'form' or 'report', not {kind!r}")
    if len(clsid) != GUID_LENGTH:
        raise AccessError(f"a TypeInfo CLSID is {GUID_LENGTH} bytes")
    out = bytearray(TYPE_INFO_MAGIC)
    out += struct.pack("<IiI", TYPE_INFO_KIND[kind], -1, len(entries))
    out += clsid
    for entry in entries:
        out += struct.pack("<II", entry.ident, entry.ordinal)
        out += entry.name.encode("cp1252", errors="replace") + bytes(2) + entry.tail
    return bytes(out)

