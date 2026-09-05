"""The metadata section of a DataMashup blob.

The section is a small envelope -- a version, the XML, and a content
package -- around a ``LocalPackageMetadataFile`` document that lists one
item per formula.  An item with entries is a query; an item without them
is one of the query's steps.  That is the rule Excel goes by: a query
whose item carries an empty ``StableEntries`` does not appear in the
Queries pane at all (measured against live Excel).

Values carry their type in the first character of the string: ``l`` for
an integer, ``f`` for a double, ``s`` for text, ``c`` for a GUID and
``d`` for a timestamp.  Those five come from Microsoft's own
``SerializedMetadataEntry``, as does the whole of :func:`pack_groups`.
"""

from __future__ import annotations

import datetime as _dt
import struct
import uuid
from dataclasses import dataclass, field
from xml.etree import ElementTree

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery._binary import BinaryReader, BinaryWriter

#: The document element and the namespaces Excel writes on it.
_ROOT = "LocalPackageMetadataFile"
_XSD = "http://www.w3.org/2001/XMLSchema"
_XSI = "http://www.w3.org/2001/XMLSchema-instance"
#: The XML declaration, byte for byte as Excel writes it, BOM included.
_DECLARATION = '﻿<?xml version="1.0" encoding="utf-8"?>'
#: An empty OPC package: the content slot of a metadata section that has
#: nothing cached in it.  Every workbook measured carries exactly this.
EMPTY_CONTENT = bytes.fromhex("504b0506" + "00" * 18)
#: The section's own version word.
METADATA_VERSION = 0

#: What each value prefix means.
KIND_LONG = "l"
KIND_DOUBLE = "f"
KIND_TEXT = "s"
KIND_GUID = "c"
KIND_TIME = "d"
_KINDS = (KIND_LONG, KIND_DOUBLE, KIND_TEXT, KIND_GUID, KIND_TIME)

#: The item types the format defines, from ``SerializedPackageItemType``.
ITEM_FORMULA = "Formula"
ITEM_ALL_FORMULAS = "AllFormulas"
ITEM_EMBEDDED_FORMULA = "EmbeddedFormula"

#: The section every Excel query lives in.
SECTION = "Section1"

#: Entry keys, spelled as Microsoft's ``QueryMetadataKey`` spells them.
IS_PRIVATE = "IsPrivate"
FILL_ENABLED = "FillEnabled"
FILL_OBJECT_TYPE = "FillObjectType"
FILL_TO_DATA_MODEL_ENABLED = "FillToDataModelEnabled"
FILL_TARGET = "FillTarget"
FILL_TARGET_NAME_CUSTOMIZED = "FillTargetNameCustomized"
FILL_LAST_UPDATED = "FillLastUpdated"
FILL_COLUMN_NAMES = "FillColumnNames"
FILL_COLUMN_TYPES = "FillColumnTypes"
FILL_STATUS = "FillStatus"
FILL_COUNT = "FillCount"
FILL_ERROR_CODE = "FillErrorCode"
FILL_ERROR_COUNT = "FillErrorCount"
FILL_ERROR_MESSAGE = "FillErrorMessage"
BUFFER_NEXT_REFRESH = "BufferNextRefresh"
NAME_UPDATED_AFTER_FILL = "NameUpdatedAfterFill"
RESULT_TYPE = "ResultType"
IS_FUNCTION_QUERY = "IsFunctionQuery"
IS_HIDDEN = "IsHidden"
QUERY_ID = "QueryID"
QUERY_GROUP_ID = "QueryGroupID"
LOADED_TO_ANALYSIS_SERVICES = "LoadedToAnalysisServices"
ADDED_TO_DATA_MODEL = "AddedToDataModel"
RELATIONSHIP_INFO_CONTAINER = "RelationshipInfoContainer"
#: Keys that belong to the document rather than to one query.
QUERY_GROUPS = "QueryGroups"
RELATIONSHIPS = "Relationships"
IS_RELATIONSHIP_DETECTION_ENABLED = "IsRelationshipDetectionEnabled"
IS_TYPE_DETECTION_ENABLED = "IsTypeDetectionEnabled"

#: ``FillObjectType`` values.
FILL_CONNECTION_ONLY = "ConnectionOnly"
FILL_TABLE = "Table"
FILL_PIVOT_TABLE = "PivotTable"
FILL_PIVOT_CHART = "PivotChart"

#: The characters .NET's ``Uri.EscapeDataString`` leaves alone.  Measured
#: over every printable ASCII character through Microsoft's own
#: ``ItemPathFromParts``.
_UNRESERVED = frozenset(
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~!'()*"
)


def format_double(value: float) -> str:
    """A double, spelled the way the writer of this format spells it.

    That writer is .NET Framework's ``double.ToString()`` under the
    invariant culture, which is the ``G15`` format: fifteen significant
    digits, an exponent once it leaves the range ``G`` keeps in plain
    form, and no sign on a negative zero.  Measured against Microsoft's
    ``SerializedMetadataEntry`` over the range in
    ``tests/test_powerquery_metadata.py``.
    """
    if value != value:  # NaN
        return "NaN"
    if value == float("inf"):
        return "Infinity"
    if value == float("-inf"):
        return "-Infinity"
    if value == 0:
        return "0"
    return f"{value:.15G}"


def escape_path_part(part: str) -> str:
    """One item-path part, escaped the way the format escapes it."""
    out: list[str] = []
    for char in part:
        if char in _UNRESERVED:
            out.append(char)
        else:
            out.extend(f"%{byte:02X}" for byte in char.encode("utf-8"))
    return "".join(out)


def unescape_path_part(part: str) -> str:
    """One item-path part, read back."""
    out = bytearray()
    at = 0
    while at < len(part):
        char = part[at]
        if char == "%" and at + 2 < len(part) + 1:
            try:
                out.append(int(part[at + 1 : at + 3], 16))
            except ValueError:
                raise PowerQueryError(f"{part!r} has a broken percent escape") from None
            at += 3
            continue
        out += char.encode("utf-8")
        at += 1
    return out.decode("utf-8")


def item_path(*parts: str) -> str:
    """The item path for these parts, escaped part by part."""
    return "/".join(escape_path_part(part) for part in parts)


def path_parts(path: str) -> tuple[str, ...]:
    """The parts an item path is made of."""
    if not path:
        return ()
    return tuple(unescape_path_part(part) for part in path.split("/"))


@dataclass(frozen=True)
class Entry:
    """One ``StableEntries`` record: a key and a typed value."""

    key: str
    raw: str

    @property
    def kind(self) -> str:
        return self.raw[:1]

    @property
    def value(self) -> object:
        """The value as Python sees it, or the raw string when the prefix
        is one this format does not define."""
        kind, body = self.raw[:1], self.raw[1:]
        if kind == KIND_LONG:
            return int(body)
        if kind == KIND_DOUBLE:
            return float(body)
        if kind == KIND_TEXT:
            return body
        if kind == KIND_GUID:
            return uuid.UUID(body)
        if kind == KIND_TIME:
            return _dt.datetime.strptime(body[:27] + "Z", "%Y-%m-%dT%H:%M:%S.%f0Z").replace(
                tzinfo=_dt.timezone.utc
            )
        return self.raw

    @property
    def flag(self) -> bool:
        """An integer entry read as the flag it stands for."""
        return self.raw == KIND_LONG + "1"

    @classmethod
    def of_int(cls, key: str, value: int) -> Entry:
        return cls(key, f"{KIND_LONG}{value}")

    @classmethod
    def of_flag(cls, key: str, value: bool) -> Entry:  # noqa: FBT001 - mirrors the wire shape
        return cls(key, f"{KIND_LONG}{1 if value else 0}")

    @classmethod
    def of_text(cls, key: str, value: str) -> Entry:
        return cls(key, f"{KIND_TEXT}{value}")

    @classmethod
    def of_double(cls, key: str, value: float) -> Entry:
        return cls(key, f"{KIND_DOUBLE}{format_double(value)}")

    @classmethod
    def of_guid(cls, key: str, value: uuid.UUID) -> Entry:
        return cls(key, f"{KIND_GUID}{value}")

    @classmethod
    def of_time(cls, key: str, value: _dt.datetime) -> Entry:
        moment = value.astimezone(_dt.timezone.utc) if value.tzinfo else value
        stamp = moment.strftime("%Y-%m-%dT%H:%M:%S.") + f"{moment.microsecond:06d}0Z"
        return cls(key, f"{KIND_TIME}{stamp}")


@dataclass
class Item:
    """One ``Item``: where it sits, and what is recorded about it."""

    item_type: str
    path: str
    entries: list[Entry] = field(default_factory=lambda: [])

    @property
    def parts(self) -> tuple[str, ...]:
        return path_parts(self.path)

    @property
    def is_query(self) -> bool:
        """Excel shows an item as a query when it is a formula directly
        under the section and carries at least one entry.  An item with no
        entries is a step (measured: an entry-less item is not listed)."""
        return self.item_type == ITEM_FORMULA and len(self.parts) == 2 and bool(self.entries)

    @property
    def is_step(self) -> bool:
        return self.item_type == ITEM_FORMULA and len(self.parts) == 3

    def get(self, key: str) -> Entry | None:
        for entry in self.entries:
            if entry.key == key:
                return entry
        return None

    def set(self, entry: Entry) -> None:
        """Replace the entry with this key, keeping its place, or append."""
        for index, existing in enumerate(self.entries):
            if existing.key == entry.key:
                self.entries[index] = entry
                return
        self.entries.append(entry)

    def remove(self, key: str) -> None:
        self.entries = [entry for entry in self.entries if entry.key != key]


@dataclass
class QueryGroup:
    """A folder in the Queries pane."""

    id: uuid.UUID
    name: str
    description: str = ""
    parent_id: uuid.UUID | None = None
    order: int = 0


def unpack_groups(value: str) -> list[QueryGroup]:
    """The groups a ``QueryGroups`` entry value holds."""
    import base64

    reader = BinaryReader(base64.b64decode(value))
    count = reader.uint32()
    groups: list[QueryGroup] = []
    for _ in range(count):
        reader.uint32()  # a word Microsoft's writer leaves at zero
        identifier = reader.guid()
        name = reader.text()
        description = reader.text()
        parent = reader.guid() if reader.boolean() else None
        groups.append(QueryGroup(identifier, name, description, parent, reader.int32()))
    return groups


def pack_groups(groups: list[QueryGroup]) -> str:
    """The ``QueryGroups`` entry value for these groups."""
    import base64

    writer = BinaryWriter()
    writer.uint32(len(groups))
    for group in groups:
        writer.uint32(0)
        writer.guid(group.id)
        writer.text(group.name)
        writer.text(group.description or "")
        writer.boolean(group.parent_id is not None)
        if group.parent_id is not None:
            writer.guid(group.parent_id)
        writer.int32(group.order)
    return base64.b64encode(writer.bytes()).decode("ascii")


def _escape_attribute(value: str) -> str:
    """An attribute value, escaped as .NET's XML writer escapes it."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("\r", "&#xD;")
        .replace("\n", "&#xA;")
        .replace("\t", "&#x9;")
    )


def _escape_text(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#xD;")


@dataclass
class Metadata:
    """A parsed metadata section: its items and its content package."""

    items: list[Item] = field(default_factory=lambda: [])
    content: bytes = EMPTY_CONTENT
    version: int = METADATA_VERSION

    # -- reading ------------------------------------------------------------

    @classmethod
    def parse(cls, raw: bytes) -> Metadata:
        if len(raw) < 8:
            raise PowerQueryError("a metadata section is at least eight bytes")
        version, length = struct.unpack_from("<II", raw, 0)
        if 8 + length > len(raw):
            raise PowerQueryError(f"the metadata XML claims {length} bytes and the section is shorter")
        xml = raw[8 : 8 + length]
        rest = raw[8 + length :]
        if len(rest) < 4:
            raise PowerQueryError("a metadata section carries a content package")
        content_length = struct.unpack_from("<I", rest, 0)[0]
        if 4 + content_length > len(rest):
            raise PowerQueryError(
                f"the metadata content claims {content_length} bytes and the section is shorter"
            )
        return cls(items=parse_items(xml), content=rest[4 : 4 + content_length], version=version)

    def serialize(self) -> bytes:
        xml = self.to_xml().encode("utf-8")
        return (
            struct.pack("<II", self.version, len(xml))
            + xml
            + struct.pack("<I", len(self.content))
            + self.content
        )

    # -- the document ------------------------------------------------------

    def to_xml(self) -> str:
        """The document, spelled the way Excel spells it: no line breaks,
        a space before every empty element's slash."""
        out = [
            _DECLARATION,
            f'<{_ROOT} xmlns:xsd="{_XSD}" xmlns:xsi="{_XSI}"><Items>',
        ]
        for item in self.items:
            out.append("<Item><ItemLocation>")
            out.append(f"<ItemType>{_escape_text(item.item_type)}</ItemType>")
            out.append(
                f"<ItemPath>{_escape_text(item.path)}</ItemPath>" if item.path else "<ItemPath />"
            )
            out.append("</ItemLocation>")
            if item.entries:
                out.append("<StableEntries>")
                for entry in item.entries:
                    out.append(
                        f'<Entry Type="{_escape_attribute(entry.key)}"'
                        f' Value="{_escape_attribute(entry.raw)}" />'
                    )
                out.append("</StableEntries>")
            else:
                out.append("<StableEntries />")
            out.append("</Item>")
        out.append(f"</Items></{_ROOT}>")
        return "".join(out)

    # -- the queries in it -------------------------------------------------

    def queries(self) -> list[Item]:
        return [item for item in self.items if item.is_query]

    def query(self, name: str) -> Item | None:
        for item in self.items:
            if item.is_query and item.parts[1] == name:
                return item
        return None

    def all_formulas(self) -> Item:
        """The document-wide item, created when it is missing."""
        for item in self.items:
            if item.item_type == ITEM_ALL_FORMULAS:
                return item
        item = Item(ITEM_ALL_FORMULAS, "")
        self.items.insert(0, item)
        return item

    def groups(self) -> list[QueryGroup]:
        entry = self.all_formulas().get(QUERY_GROUPS)
        if entry is None:
            return []
        return unpack_groups(entry.raw[1:])

    def set_groups(self, groups: list[QueryGroup]) -> None:
        item = self.all_formulas()
        if groups:
            item.set(Entry.of_text(QUERY_GROUPS, pack_groups(groups)))
        else:
            item.remove(QUERY_GROUPS)

    def steps_of(self, name: str) -> list[Item]:
        return [item for item in self.items if item.is_step and item.parts[1] == name]

    def index_of(self, item: Item) -> int:
        for index, candidate in enumerate(self.items):
            if candidate is item:
                return index
        raise PowerQueryError("that item does not belong to this metadata")

    def drop_query(self, name: str) -> None:
        """The query's item and its steps, gone."""
        self.items = [
            item
            for item in self.items
            if not (item.item_type == ITEM_FORMULA and item.parts[1:2] == (name,))
        ]

    def rename_query(self, old: str, new: str) -> None:
        for item in self.items:
            if item.item_type == ITEM_FORMULA and item.parts[1:2] == (old,):
                parts = list(item.parts)
                parts[1] = new
                item.path = item_path(*parts)

    def set_steps(self, name: str, steps: list[str]) -> None:
        """Rewrite the query's step items, in order, right after it."""
        query = self.query(name)
        if query is None:
            raise PowerQueryError(f"this metadata has no query named {name!r}")
        kept = [
            item
            for item in self.items
            if not (item.is_step and item.parts[1] == name)
        ]
        at = kept.index(query) + 1
        fresh = [Item(ITEM_FORMULA, item_path(SECTION, name, step)) for step in steps]
        self.items = kept[:at] + fresh + kept[at:]


def parse_items(xml: bytes) -> list[Item]:
    """The items a ``LocalPackageMetadataFile`` document lists."""
    try:
        root = ElementTree.fromstring(xml.decode("utf-8-sig"))
    except ElementTree.ParseError as exc:
        raise PowerQueryError(f"the metadata XML does not parse: {exc}") from exc
    if root.tag.rsplit("}", 1)[-1] != _ROOT:
        raise PowerQueryError(f"the metadata document is <{root.tag}>, not <{_ROOT}>")
    items: list[Item] = []
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "Item":
            continue
        location = _child(element, "ItemLocation")
        if location is None:
            raise PowerQueryError("an item carries no ItemLocation")
        item_type = _text(_child(location, "ItemType"))
        path = _text(_child(location, "ItemPath"))
        entries: list[Entry] = []
        stable = _child(element, "StableEntries")
        if stable is not None:
            for entry in stable:
                if entry.tag.rsplit("}", 1)[-1] != "Entry":
                    continue
                entries.append(Entry(entry.get("Type", ""), entry.get("Value", "")))
        items.append(Item(item_type, path, entries))
    return items


def _child(element: ElementTree.Element, name: str) -> ElementTree.Element | None:
    for candidate in element:
        if candidate.tag.rsplit("}", 1)[-1] == name:
            return candidate
    return None


def _text(element: ElementTree.Element | None) -> str:
    if element is None or element.text is None:
        return ""
    return element.text
