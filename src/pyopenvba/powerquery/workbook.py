"""Power Query in an Excel workbook: the queries, and the file around them.

A workbook keeps its whole Power Query project in one custom XML part, as
a base64 blob.  :class:`PowerQueryWorkbook` opens the workbook, reads the
blob, and hands back the queries; changing one and saving puts the blob
back with everything else in the file untouched.

    from pyopenvba import PowerQueryWorkbook

    with PowerQueryWorkbook("orders.xlsx") as book:
        print(book.query_names())
        book.query("Orders").formula = "let Source = Excel.CurrentWorkbook() in Source"
        book.add_query("Totals", "let Source = Table.RowCount(Orders) in Source")
        book.save()

Adding a query means two things at once: a ``shared`` member in the
section document and an item in the metadata.  Excel needs both -- a
member with no metadata item does not appear in the Queries pane, and an
item with no member makes the workbook fail to open its queries.  Both
were measured; see :mod:`pyopenvba.powerquery._metadata`.
"""

from __future__ import annotations

import base64
import re
import uuid
from pathlib import Path
from typing import ClassVar
from xml.etree import ElementTree

from pyopenvba.exceptions import PowerQueryError
from pyopenvba.powerquery import _metadata as meta
from pyopenvba.powerquery._mashup import DEFAULT_PERMISSIONS, Mashup
from pyopenvba.powerquery._metadata import Entry, Item, Metadata, QueryGroup
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._package import SECTION_PART, new_package
from pyopenvba.powerquery._section import Section, is_function_expression, new_section
from pyopenvba.powerquery._sheets import load_to_sheet as _load_to_sheet
from pyopenvba.powerquery._sheets import unload_from_sheet as _unload_from_sheet

#: The namespace that marks the custom XML part as a mashup.
DATA_MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
#: The wrapper Excel writes around the base64, in UTF-16 with a BOM.
_ITEM_HEAD = '<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="' + DATA_MASHUP_NS + '">'
_ITEM_TAIL = "</DataMashup>"
_BOM = b"\xff\xfe"

_CUSTOM_XML_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
)
_CUSTOM_XML_PROPS_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
)
_CUSTOM_XML_PROPS_TYPE = (
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
)
_WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
_CONTENT_TYPES = "[Content_Types].xml"
_ITEM_PATTERN = re.compile(r"^customXml/item(\d+)\.xml$")

#: What a query's result is loaded into.
LOAD_CONNECTION_ONLY = "connection-only"
LOAD_TABLE = "table"
LOAD_PIVOT_TABLE = "pivot-table"
_FILL_TYPES = {
    LOAD_CONNECTION_ONLY: meta.FILL_CONNECTION_ONLY,
    LOAD_TABLE: meta.FILL_TABLE,
    LOAD_PIVOT_TABLE: meta.FILL_PIVOT_TABLE,
}
_LOAD_NAMES = {value: key for key, value in _FILL_TYPES.items()}


def _wrap(blob: bytes, head: str = _ITEM_HEAD, tail: str = _ITEM_TAIL) -> bytes:
    """The custom XML part for this blob, inside the wrapper it came in.

    The wrapper is kept rather than rebuilt because Excel sometimes puts
    an ``sqmid`` attribute on the element, and a workbook that carries one
    should still write back byte for byte.
    """
    text = head + base64.b64encode(blob).decode("ascii") + tail
    return _BOM + text.encode("utf-16-le")


def _split_item(raw: bytes) -> tuple[str, bytes, str]:
    """The part's wrapper and the blob inside it."""
    text = raw.decode("utf-16") if raw[:2] == _BOM else raw.decode("utf-8-sig")
    match = re.search(r"(.*<DataMashup[^>]*>)(.*?)(</DataMashup>.*)", text, re.S)
    if match is None:
        raise PowerQueryError("this custom XML part carries no DataMashup element")
    try:
        return match.group(1), base64.b64decode(match.group(2), validate=False), match.group(3)
    except (ValueError, TypeError) as exc:
        raise PowerQueryError(f"the DataMashup base64 does not decode: {exc}") from exc


def _is_mashup_part(raw: bytes) -> bool:
    if b"DataMashup" not in raw and b"D\x00a\x00t\x00a\x00M\x00a\x00s\x00h\x00u\x00p" not in raw:
        return False
    try:
        text = raw.decode("utf-16") if raw[:2] == _BOM else raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return False
    return f"{{{DATA_MASHUP_NS}}}" in _root_tag(text) if text.lstrip().startswith("<") else False


def _root_tag(text: str) -> str:
    try:
        return ElementTree.fromstring(text).tag
    except ElementTree.ParseError:
        return ""


class PowerQuery:
    """One query: its M, and what the workbook records about it."""

    def __init__(self, book: PowerQueryWorkbook, name: str) -> None:
        self._book = book
        self._name = name

    # -- identity -----------------------------------------------------------

    @property
    def name(self) -> str:
        return self._name

    def rename(self, name: str, *, update_references: bool = True) -> None:
        self._book.rename_query(self._name, name, update_references=update_references)
        self._name = name

    # -- the M --------------------------------------------------------------

    @property
    def formula(self) -> str:
        return self._book.section.formula(self._name)

    @formula.setter
    def formula(self, formula: str) -> None:
        self._book._set_formula(self._name, formula)  # pyright: ignore[reportPrivateUsage]

    @property
    def steps(self) -> list[str]:
        """The names of the query's top-level ``let`` bindings."""
        return self._book.section.steps(self._name)

    @property
    def description(self) -> str | None:
        return self._book.section.description(self._name)

    @description.setter
    def description(self, description: str | None) -> None:
        self._book._set_description(self._name, description)  # pyright: ignore[reportPrivateUsage]

    @property
    def is_function(self) -> bool:
        return is_function_expression(self.formula)

    # -- what the metadata records -----------------------------------------

    @property
    def _item(self) -> Item:
        item = self._book.metadata.query(self._name)
        if item is None:  # pragma: no cover - the two halves are kept in step
            raise PowerQueryError(f"the query {self._name!r} has no metadata item")
        return item

    def entries(self) -> dict[str, object]:
        """Everything the metadata records about this query, decoded."""
        return {entry.key: entry.value for entry in self._item.entries}

    def entry(self, key: str) -> Entry | None:
        return self._item.get(key)

    def set_entry(self, entry: Entry) -> None:
        self._item.set(entry)
        self._book._touch()  # pyright: ignore[reportPrivateUsage]

    @property
    def query_id(self) -> str | None:
        entry = self._item.get(meta.QUERY_ID)
        return None if entry is None else str(entry.value)

    @property
    def is_private(self) -> bool:
        entry = self._item.get(meta.IS_PRIVATE)
        return bool(entry and entry.flag)

    @property
    def load_target(self) -> str:
        """Where the query's result goes: ``connection-only``, ``table`` or
        ``pivot-table``."""
        entry = self._item.get(meta.FILL_OBJECT_TYPE)
        if entry is None:
            return LOAD_CONNECTION_ONLY
        return _LOAD_NAMES.get(str(entry.value), str(entry.value))

    @property
    def loads_to_data_model(self) -> bool:
        entry = self._item.get(meta.FILL_TO_DATA_MODEL_ENABLED)
        return bool(entry and entry.flag)

    @property
    def load_enabled(self) -> bool:
        entry = self._item.get(meta.FILL_ENABLED)
        return bool(entry and entry.flag)

    @property
    def target_name(self) -> str | None:
        """The worksheet table a loaded query fills."""
        entry = self._item.get(meta.FILL_TARGET)
        return None if entry is None else str(entry.value)

    @property
    def group(self) -> QueryGroup | None:
        entry = self._item.get(meta.QUERY_GROUP_ID)
        if entry is None:
            return None
        wanted = str(entry.value)
        for group in self._book.groups():
            if str(group.id) == wanted:
                return group
        return None

    def load_to_sheet(
        self,
        columns: list[str],
        *,
        sheet: str | int = 1,
        cell: str = "A1",
        table_name: str | None = None,
    ) -> str:
        """Put this query's result on a worksheet, in a table of its own."""
        return self._book.load_to_sheet(
            self._name, columns, sheet=sheet, cell=cell, table_name=table_name
        )

    def unload(self) -> bool:
        """Take the query's table off its sheet and leave it a connection."""
        return self._book.unload(self._name)

    def move_to_group(self, group: QueryGroup | None) -> None:
        """Put the query in a group, or take it out of the one it is in."""
        if group is None:
            self._item.remove(meta.QUERY_GROUP_ID)
        else:
            self._item.set(Entry.of_text(meta.QUERY_GROUP_ID, str(group.id)))
        self._book._touch()  # pyright: ignore[reportPrivateUsage]

    def __repr__(self) -> str:
        return f"PowerQuery({self._name!r}, {self.load_target})"


class PowerQueryWorkbook:
    """The Power Query project of one workbook."""

    #: The Excel formats that carry a package, and so can carry queries.
    formats: ClassVar[frozenset[str]] = frozenset({".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlam"})

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        if self.path.suffix.lower() not in self.formats:
            raise PowerQueryError(
                f"{self.path.suffix or 'this file'} is not an Excel package; "
                f"expected one of {', '.join(sorted(self.formats))}"
            )
        self._opc = OpcFile.parse(self.path.read_bytes())
        self._part = self._find_part()
        self._mashup: Mashup | None = None
        self._head, self._tail = _ITEM_HEAD, _ITEM_TAIL
        if self._part is not None:
            self._head, blob, self._tail = _split_item(self._opc.read(self._part))
            self._mashup = Mashup.parse(blob)
            self._section = Section(self._mashup.package.read(SECTION_PART).decode("utf-8"))
        else:
            self._section = new_section()

    # -- opening and closing ------------------------------------------------

    @classmethod
    def create_new(cls, path: str | Path) -> PowerQueryWorkbook:
        """Make a workbook at `path` with nothing in it, and open it.

        The bytes come from a workbook Excel saved empty, so the file
        opens with no repair prompt; the queries added to it bring the
        custom XML part along the first time one is written.  An existing
        file at `path` is replaced.
        """
        from pyopenvba._templates import EMPTY_XLSX_BYTES

        target = Path(path)
        if target.suffix.lower() not in cls.formats:
            raise PowerQueryError(
                f"{target.suffix or 'this file'} is not an Excel package; "
                f"expected one of {', '.join(sorted(cls.formats))}"
            )
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(EMPTY_XLSX_BYTES)
        return cls(target)

    def __enter__(self) -> PowerQueryWorkbook:
        return self

    def __exit__(self, *_: object) -> None:
        return None

    def _find_part(self) -> str | None:
        for name in self._opc.names():
            if _ITEM_PATTERN.match(name) and _is_mashup_part(self._opc.read(name)):
                return name
        return None

    # -- what is in it ------------------------------------------------------

    @property
    def has_queries(self) -> bool:
        return self._mashup is not None and bool(self.metadata.queries())

    @property
    def section(self) -> Section:
        """The M section document behind the queries."""
        return self._section

    @property
    def metadata(self) -> Metadata:
        if self._mashup is None:
            raise PowerQueryError("this workbook has no Power Query package yet")
        return self._mashup.metadata

    @property
    def mashup(self) -> Mashup:
        if self._mashup is None:
            raise PowerQueryError("this workbook has no Power Query package yet")
        return self._mashup

    def section_text(self) -> str:
        """The whole section document, as it is stored."""
        return self._section.text

    def query_names(self) -> list[str]:
        """The queries, in the order the section lists them."""
        if self._mashup is None:
            return []
        known = {item.parts[1] for item in self.metadata.queries()}
        return [name for name in self._section.names() if name in known]

    def queries(self) -> list[PowerQuery]:
        return [PowerQuery(self, name) for name in self.query_names()]

    def query(self, name: str) -> PowerQuery:
        for candidate in self.query_names():
            if candidate == name:
                return PowerQuery(self, name)
        raise PowerQueryError(f"this workbook has no query named {name!r}")

    def groups(self) -> list[QueryGroup]:
        return [] if self._mashup is None else self.metadata.groups()

    # -- changing it --------------------------------------------------------

    def _touch(self) -> None:
        self.mashup.touch()

    def _sync_section(self) -> None:
        self.mashup.package.write(SECTION_PART, self._section.text.encode("utf-8"))
        self._touch()

    def _set_formula(self, name: str, formula: str) -> None:
        self._section.set_formula(name, formula)
        self.metadata.set_steps(name, self._section.steps(name))
        self._sync_section()

    def _set_description(self, name: str, description: str | None) -> None:
        self._section.set_description(name, description)
        self._sync_section()

    def _ensure_package(self) -> None:
        """Give a workbook that has never held a query the package Excel
        would have given it."""
        if self._mashup is not None:
            return
        package = new_package(self._section.text)
        document = Metadata()
        document.items.append(Item(meta.ITEM_ALL_FORMULAS, ""))
        self._mashup = Mashup(package=package, metadata=document, permissions=DEFAULT_PERMISSIONS)

    def add_query(
        self,
        name: str,
        formula: str,
        *,
        description: str | None = None,
        group: QueryGroup | None = None,
        query_id: uuid.UUID | None = None,
    ) -> PowerQuery:
        """Add a query, with the entries Excel writes for a new one."""
        if not name:
            raise PowerQueryError("a query needs a name")
        if self._section.has(name):
            raise PowerQueryError(f"this workbook already has a query named {name!r}")
        self._ensure_package()
        self._section.add(name, formula, description)
        item = Item(
            meta.ITEM_FORMULA,
            meta.item_path(meta.SECTION, name),
            [
                Entry.of_flag(meta.IS_PRIVATE, False),
                Entry.of_flag(meta.FILL_ENABLED, False),
                Entry.of_text(meta.FILL_OBJECT_TYPE, meta.FILL_CONNECTION_ONLY),
                Entry.of_flag(meta.FILL_TO_DATA_MODEL_ENABLED, False),
                Entry.of_text(meta.QUERY_ID, str(query_id or uuid.uuid4())),
            ],
        )
        if is_function_expression(formula):
            item.set(Entry.of_text(meta.RESULT_TYPE, "Function"))
        if group is not None:
            item.set(Entry.of_text(meta.QUERY_GROUP_ID, str(group.id)))
        self.metadata.items.append(item)
        self.metadata.set_steps(name, self._section.steps(name))
        self._sync_section()
        return PowerQuery(self, name)

    def remove_query(self, name: str) -> None:
        self.query(name)
        self._section.remove(name)
        self.metadata.drop_query(name)
        self._sync_section()

    def rename_query(self, old: str, new: str, *, update_references: bool = True) -> None:
        """Rename a query, and by default the references to it.

        Excel's own editor rewrites the queries that name this one; the
        rewrite here goes through the M tokenizer, so a match inside a
        text literal, a comment or a record's field name is left alone.
        """
        self.query(old)
        if self._section.has(new):
            raise PowerQueryError(f"this workbook already has a query named {new!r}")
        self._section.rename(old, new, update_references=update_references)
        self.metadata.rename_query(old, new)
        self._sync_section()

    def add_group(self, name: str, *, parent: QueryGroup | None = None, description: str = "") -> QueryGroup:
        """Add a folder to the Queries pane."""
        self._ensure_package()
        groups = self.metadata.groups()
        if any(group.name == name and group.parent_id == (parent.id if parent else None) for group in groups):
            raise PowerQueryError(f"this workbook already has a group named {name!r} there")
        order = max((group.order for group in groups), default=-1) + 1
        group = QueryGroup(uuid.uuid4(), name, description, parent.id if parent else None, order)
        groups.append(group)
        self.metadata.set_groups(groups)
        self._touch()
        return group

    def remove_group(self, group: QueryGroup) -> None:
        """Drop a group; the queries in it move back to the top level."""
        groups = [candidate for candidate in self.metadata.groups() if candidate.id != group.id]
        for child in groups:
            if child.parent_id == group.id:
                child.parent_id = group.parent_id
        self.metadata.set_groups(groups)
        for query in self.queries():
            if query.entry(meta.QUERY_GROUP_ID) is not None and str(
                query.entry(meta.QUERY_GROUP_ID).value  # pyright: ignore[reportOptionalMemberAccess]
            ) == str(group.id):
                query.move_to_group(None)
        self._touch()

    def load_to_sheet(
        self,
        query: str,
        columns: list[str],
        *,
        sheet: str | int = 1,
        cell: str = "A1",
        table_name: str | None = None,
    ) -> str:
        """Put a query's result on a worksheet, and say so in the metadata.

        `columns` names the columns the query returns.  They have to be
        given because working them out means running the query, which
        only the mashup engine can do; Excel settles them against the
        real result, and fills the rows, on its first refresh.

        Returns the name of the table that was made.
        """
        item = self.query(query)._item  # pyright: ignore[reportPrivateUsage]
        if item.get(meta.FILL_ENABLED) is not None and item.get(meta.FILL_ENABLED).flag:  # pyright: ignore[reportOptionalMemberAccess]
            raise PowerQueryError(f"the query {query!r} already loads somewhere; unload it first")
        name = _load_to_sheet(
            self._opc, query, columns, sheet=sheet, cell=cell, table_name=table_name
        )
        item.set(Entry.of_flag(meta.FILL_ENABLED, True))
        item.set(Entry.of_text(meta.FILL_OBJECT_TYPE, meta.FILL_TABLE))
        item.set(Entry.of_text(meta.FILL_TARGET, name))
        item.set(Entry.of_flag(meta.FILL_TARGET_NAME_CUSTOMIZED, True))
        item.set(Entry.of_flag(meta.NAME_UPDATED_AFTER_FILL, False))
        self._touch()
        return name

    def unload(self, query: str) -> bool:
        """Make a loaded query connection-only again."""
        item = self.query(query)._item  # pyright: ignore[reportPrivateUsage]
        removed = _unload_from_sheet(self._opc, query)
        item.set(Entry.of_flag(meta.FILL_ENABLED, False))
        item.set(Entry.of_text(meta.FILL_OBJECT_TYPE, meta.FILL_CONNECTION_ONLY))
        for key in (meta.FILL_TARGET, meta.FILL_TARGET_NAME_CUSTOMIZED, meta.NAME_UPDATED_AFTER_FILL):
            item.remove(key)
        self._touch()
        return removed

    def set_section_text(self, text: str) -> None:
        """Replace the whole section document.

        The metadata follows: a query that the new document drops loses
        its item, and one it adds gets the entries Excel gives a new
        query.  Otherwise Excel would show a stale list.
        """
        self._ensure_package()
        fresh = Section(text)
        before = set(self._section.names())
        self._section = fresh
        for gone in before - set(fresh.names()):
            self.metadata.drop_query(gone)
        for name in fresh.names():
            if self.metadata.query(name) is None:
                self.metadata.items.append(
                    Item(
                        meta.ITEM_FORMULA,
                        meta.item_path(meta.SECTION, name),
                        [
                            Entry.of_flag(meta.IS_PRIVATE, False),
                            Entry.of_flag(meta.FILL_ENABLED, False),
                            Entry.of_text(meta.FILL_OBJECT_TYPE, meta.FILL_CONNECTION_ONLY),
                            Entry.of_flag(meta.FILL_TO_DATA_MODEL_ENABLED, False),
                            Entry.of_text(meta.QUERY_ID, str(uuid.uuid4())),
                        ],
                    )
                )
            self.metadata.set_steps(name, fresh.steps(name))
        self._sync_section()

    # -- writing ------------------------------------------------------------

    def to_bytes(self) -> bytes:
        """The workbook's bytes, with the package as it stands."""
        if self._mashup is not None:
            blob = self._mashup.serialize()
            if self._part is None:
                self._part = self._install_part()
            self._opc.write(self._part, _wrap(blob, self._head, self._tail))
        return self._opc.serialize()

    def save(self, path: str | Path | None = None) -> Path:
        out = Path(path) if path is not None else self.path
        raw = self.to_bytes()
        out.write_bytes(raw)
        if path is None:
            self._opc = OpcFile.parse(raw)
        return out

    # -- making room for a package that was never there ---------------------

    def _next_item_index(self) -> int:
        used = {
            int(match.group(1))
            for match in (_ITEM_PATTERN.match(name) for name in self._opc.names())
            if match
        }
        index = 1
        while index in used:
            index += 1
        return index

    def _install_part(self) -> str:
        """Write the custom XML part and the bookkeeping around it.

        A workbook with no Power Query at all needs four things beyond the
        blob: the properties part beside it, a relationship from the part
        to those properties, a content type for them, and a relationship
        from the workbook to the part.  Measured against workbooks Excel
        wrote: nothing else is required for queries that only hold a
        connection.
        """
        index = self._next_item_index()
        item = f"customXml/item{index}.xml"
        props = f"customXml/itemProps{index}.xml"
        rels = f"customXml/_rels/item{index}.xml.rels"
        self._opc.write(
            props,
            (
                '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n'
                f'<ds:datastoreItem ds:itemID="{{{str(uuid.uuid4()).upper()}}}"'
                ' xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
                f'<ds:schemaRefs><ds:schemaRef ds:uri="{DATA_MASHUP_NS}"/></ds:schemaRefs>'
                "</ds:datastoreItem>"
            ).encode("utf-8"),
        )
        self._opc.write(
            rels,
            (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{_CUSTOM_XML_PROPS_RELATIONSHIP}"'
                f' Target="itemProps{index}.xml"/></Relationships>'
            ).encode("utf-8"),
        )
        self._add_content_type(props)
        self._add_workbook_relationship(f"../customXml/item{index}.xml")
        self._opc.write(item, b"")
        return item

    def _add_content_type(self, part: str) -> None:
        raw = self._opc.read(_CONTENT_TYPES).decode("utf-8")
        override = f'<Override PartName="/{part}" ContentType="{_CUSTOM_XML_PROPS_TYPE}"/>'
        if override in raw:
            return
        if "</Types>" not in raw:
            raise PowerQueryError("the content-types part has no </Types>")
        self._opc.write(_CONTENT_TYPES, raw.replace("</Types>", override + "</Types>").encode("utf-8"))

    def _add_workbook_relationship(self, target: str) -> None:
        if not self._opc.has(_WORKBOOK_RELS):
            raise PowerQueryError("this workbook has no relationships part")
        raw = self._opc.read(_WORKBOOK_RELS).decode("utf-8")
        used = set(re.findall(r'Id="([^"]+)"', raw))
        index = 1
        while f"rId{index}" in used:
            index += 1
        relationship = (
            f'<Relationship Id="rId{index}" Type="{_CUSTOM_XML_RELATIONSHIP}" Target="{target}"/>'
        )
        if "</Relationships>" not in raw:
            raise PowerQueryError("the workbook relationships part has no </Relationships>")
        self._opc.write(
            _WORKBOOK_RELS,
            raw.replace("</Relationships>", relationship + "</Relationships>").encode("utf-8"),
        )

    def __repr__(self) -> str:
        return f"PowerQueryWorkbook({self.path.name!r}, {len(self.query_names())} queries)"
