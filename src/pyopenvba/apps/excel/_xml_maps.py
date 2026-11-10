"""XML maps made from a schema, as live Excel has them (scripts/measure_xml_maps.py, tests/fixtures/xml_maps.json).

XmlMaps.Add reads an .xsd and keeps its text as given; the map is named after its root element and _Map, a number after it
for the second of a name. A cell maps to one leaf of the schema, a single mapping, which of a repeating element is its first
item; a table's column maps to a leaf of the repeating element, and a block of one column mapped at once becomes a table.
ImportXml fills single mappings and table rows typed by the schema, and grows the table; ExportXml writes the mapped cells
back as tab-indented XML. A save writes xl/xmlMaps.xml, the tables as tableType="xml" with each column's xmlColumnPr, and a
sheet's single mappings in its tableSingleCells part, and an open reads them back.

Only flat schemas are built: a root of leaves, or of leaves and one repeating element of leaves and attributes. Anything
else a schema says, choices, references, named types, a second repeating level, reports itself. Every file a map reads or
writes goes through the sandbox gate (Application.outside) and is opened at its final_path."""
from __future__ import annotations

import re
import xml.etree.ElementTree as ElementTree
from dataclasses import dataclass, field
from decimal import Decimal
from typing import TYPE_CHECKING, Any, cast

from pyopenvba._a1 import Area
from pyopenvba._xml import escape
from pyopenvba.apps.excel._creator import ItemDefault
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._files import final_path, native_path
from pyopenvba.interpreter._objects import VBACollection, VBAObject, member, method, setter
from pyopenvba.interpreter._runtime import Slot
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, VBADate, VBAInt, error, to_bool, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._tables import Table, TableColumn

XS = "{http://www.w3.org/2001/XMLSchema}"
#: The errors Excel gives (tests/fixtures/xml_maps.json): a root the schema has no top-level element for, a schema file that is
#: not there, and anything else Excel's XML layer refuses -- an XPath it cannot find, data that is not XML.
_NO_ROOT = -2147217373
_NO_FILE = -2147217376
_UNSPECIFIED = -2147467259
#: The leaf types an import reads and an export writes, as the parts spell them.
_TYPES = {"integer", "string", "date", "boolean", "decimal"}


@dataclass
class Leaf:
    """One value a schema holds: an element of simple type, or an attribute."""
    name: str
    #: Its path with the steps' prefixes, as XPath.Value reads it: /list/item/n, /ns1:list/ns1:item/ns1:n, /list/item/@id.
    path: str
    kind: str
    #: Whether it lies under the schema's repeating element, and whether it repeats itself too.
    repeating: bool
    repeats_again: bool = False
    #: Whether the schema requires it: an element unless minOccurs is 0, an attribute where use is required.
    required: bool = True


@dataclass
class Schema:
    id: str
    #: The schema's text as its file has it, its XML declaration left out.
    xml: str
    namespace: str
    prefix: str
    root: str
    leaves: list[Leaf]
    #: The repeating element's path, /list/item, or "" for a root of leaves alone.
    repeat_path: str
    #: The names along it, root first.
    repeat_steps: list[str] = field(default_factory=lambda: list[str]())
    #: Whether the schema requires one of the repeating element at least.
    repeat_required: bool = False


@dataclass
class MapState:
    id: int
    name: str
    schema: Schema
    show_errors: bool = False
    save_source: bool = False
    autofit: bool = True
    preserve_filter: bool = True
    preserve_format: bool = True
    append: bool = False
    source_url: str = ""


@dataclass
class SingleMapping:
    map: MapState
    leaf: Leaf
    uid: str = ""
    cell_uid: str = ""


# --- reading a schema --------------------------------------------------------------------------------------------------


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _declaration_left_out(text: str) -> str:
    found = re.match(r"\s*<\?xml[^>]*\?>", text)
    return text[found.end():] if found else text


def read_schema(text: str, root: object, number: int) -> Schema:
    """A flat schema's leaves under ``root``, or the first top-level element when it is left out."""
    try:
        tree = ElementTree.fromstring(text)
    except ElementTree.ParseError:
        raise error(_UNSPECIFIED, "Automation error\nUnspecified error") from None
    if tree.tag != XS + "schema":
        raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
    for child in tree:
        if _local(child.tag) not in ("element", "annotation"):
            raise VBAUnsupportedError(f"an XML schema with a top-level {_local(child.tag)} is not implemented")
    tops = [child for child in tree if child.tag == XS + "element"]
    wanted = to_text(root) if root is not MISSING else (tops[0].get("name", "") if tops else "")
    top = next((one for one in tops if one.get("name") == wanted), None)
    if top is None:
        raise error(_NO_ROOT, "Automation error")
    namespace = tree.get("targetNamespace", "")
    qualified = tree.get("elementFormDefault") == "qualified"
    prefix = "ns1" if namespace else ""
    schema = Schema(id=f"Schema{number}", xml=_declaration_left_out(text), namespace=namespace, prefix=prefix, root=wanted,
                    leaves=[], repeat_path="")

    def step(name: str, top_level: bool) -> str:
        return f"{prefix}:{name}" if prefix and (top_level or qualified) else name

    def walk(element: ElementTree.Element, path: str, repeating: bool, depth: int) -> None:
        name = element.get("name", "")
        if not name or element.get("ref") is not None:
            raise VBAUnsupportedError("an XML schema element by reference is not implemented")
        here = path + "/" + step(name, depth == 0)
        repeats = element.get("maxOccurs", "1") not in ("1", "0")
        kind = element.get("type")
        complex_type = element.find(XS + "complexType")
        if complex_type is None:
            local = (kind or "xs:string").split(":")[-1]
            if kind is not None and not kind.startswith(("xs:", "xsd:")):
                raise VBAUnsupportedError("an XML schema element of a named type is not implemented")
            schema.leaves.append(Leaf(name, here, local, repeating or repeats, repeats_again=repeating and repeats,
                                      required=element.get("minOccurs", "1") != "0"))
            return
        if kind is not None:
            raise VBAUnsupportedError("an XML schema element of a named type is not implemented")
        if repeats:
            if repeating or schema.repeat_path:
                raise VBAUnsupportedError("an XML schema with more than one repeating element is not implemented")
            schema.repeat_path = here
            schema.repeat_steps = [one for one in here.split("/") if one]
            schema.repeat_required = element.get("minOccurs", "1") != "0"
        for part in complex_type:
            if _local(part.tag) == "sequence":
                for child in part:
                    if child.tag != XS + "element":
                        raise VBAUnsupportedError(f"an XML schema with {_local(child.tag)} in a sequence is not implemented")
                    walk(child, here, repeating or repeats, depth + 1)
            elif _local(part.tag) == "attribute":
                local = part.get("type", "xs:string").split(":")[-1]
                schema.leaves.append(Leaf(part.get("name", ""), here + "/@" + part.get("name", ""), local, repeating or repeats,
                                          required=part.get("use") == "required"))
            else:
                raise VBAUnsupportedError(f"an XML schema with {_local(part.tag)} is not implemented")

    walk(top, "", False, 0)
    return schema


def _leaf_at(schema: Schema, xpath: str) -> Leaf:
    """The leaf an XPath names, its steps matched by their names with or without the map's prefix."""
    def bare(path: str) -> list[str]:
        return [one.split(":")[-1] for one in path.split("/") if one]

    wanted = bare(xpath)
    found = next((leaf for leaf in schema.leaves if bare(leaf.path) == wanted), None)
    if found is None:
        raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
    return found


# --- the workbook's maps -------------------------------------------------------------------------------------------------


def maps_of(book: Workbook) -> list[MapState]:
    return cast("list[MapState]", book.xml_maps)


def _read_file(book: Workbook, path: str, what: str) -> str:
    """A file a macro names, read through the sandbox's gate at the path that gate checked."""
    book.application.outside(what)
    final = final_path(native_path(path))
    try:
        with open(final, "rb") as handle:
            data = handle.read()
    except OSError:
        raise error(_NO_FILE, "Automation error") from None
    return data.decode("utf-8-sig", errors="replace")


def _write_file(book: Workbook, path: str, text: str, what: str) -> None:
    book.application.outside(what)
    final = final_path(native_path(path))
    try:
        with open(final, "wb") as handle:
            handle.write(text.encode("utf-8"))
    except OSError:
        raise error(1004, f"{what} could not write the file") from None


def add_map(book: Workbook, schema_path: object, root: object) -> MapState:
    if schema_path is MISSING:
        raise error(449, "Argument not optional")
    text = _read_file(book, to_text(schema_path), "XmlMaps.Add")
    book.xml_schemas_made += 1
    schema = read_schema(text, root, book.xml_schemas_made)
    names = {one.name.casefold() for one in maps_of(book)}
    name, number = f"{schema.root}_Map", 0
    while name.casefold() in names:
        number += 1
        name = f"{schema.root}_Map{number}"
    book.xml_maps_made += 1
    made = MapState(id=book.xml_maps_made, name=name, schema=schema)
    maps_of(book).append(made)
    book.xml_maps_changed = True
    return made


def mappings(book: Workbook, state: MapState) -> tuple[list[tuple[Worksheet, tuple[int, int], SingleMapping]], list[tuple[Table, TableColumn]]]:
    singles = [(sheet, at, one) for sheet in book.sheets_ for at, one in sheet.xml_cells.items() if one.map is state]
    columns = [(table, column) for sheet in book.sheets_ for table in sheet.tables for column in table.columns
               if column.xml is not None and column.xml[0] is state]
    return singles, columns


# --- importing -----------------------------------------------------------------------------------------------------------


def _typed(text: str, kind: str) -> tuple[object, bool]:
    """A value as the schema types it, and whether the schema takes it: one it does not is kept as text."""

    if kind == "string":
        return text, True
    if kind in ("integer", "decimal"):
        pattern = r"[+-]?\d+" if kind == "integer" else r"[+-]?(\d+(\.\d*)?|\.\d+)"
        if re.fullmatch(pattern, text.strip()):
            return float(text), True
        return text, False
    if kind == "boolean":
        # true and false are Booleans, but 1 and 0 the numbers they spell (tests/fixtures/xml_maps.json).
        if text.strip() in ("true", "false"):
            return text.strip() == "true", True
        if text.strip() in ("1", "0"):
            return float(text.strip()), True
        return text, False
    if kind == "date":
        found = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", text.strip())
        if found:
            import datetime as dt

            try:
                return VBADate.from_datetime(dt.date(int(found.group(1)), int(found.group(2)), int(found.group(3)))), True
            except ValueError:
                # A day no calendar has, 2024-13-01: kept as its text, as any value the schema refuses.
                return text, False
        return text, False
    raise VBAUnsupportedError(f"importing an XML value of type {kind} is not implemented")


def _steps(path: str) -> list[str]:
    return [one.split(":")[-1] for one in path.split("/") if one]


def _children(node: ElementTree.Element, name: str) -> list[ElementTree.Element]:
    return [one for one in node if _local(one.tag) == name]


def _value_under(node: ElementTree.Element, steps: list[str]) -> str | None:
    """The text at ``steps`` below ``node``, the first where several match; an attribute is a step starting with @."""
    for place, one in enumerate(steps):
        if one.startswith("@"):
            return node.get(one[1:]) if place == len(steps) - 1 else None
        found = _children(node, one)
        if not found:
            return None
        node = found[0]
    return node.text or ""


def _parsed(data: str) -> ElementTree.Element:
    try:
        return ElementTree.fromstring(data)
    except ElementTree.ParseError:
        raise error(_UNSPECIFIED, "Automation error\nUnspecified error") from None


def keeps_empty_text(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether "" written to a cell stays a zero-length string: in the body of a table column an XML map fills it does, where
    a plain cell or plain table cell is emptied (tests/fixtures/xml_maps.json). A cell mapped on its own is not measured."""
    for table in sheet.tables:
        data = table.data
        if data is not None and data.top <= row <= data.bottom and table.area.left <= column <= table.area.right:
            return table.columns[column - table.area.left].xml is not None
    return False


def import_xml(book: Workbook, state: MapState, data: str, overwrite: bool) -> int:
    """XmlMap.ImportXml: 0, or 2 where the data's root is not the map's or a value is not its type; data that is not XML fails."""
    tree = _parsed(data)
    schema = state.schema
    if _local(tree.tag) != schema.root:
        return 2
    singles, columns = mappings(book, state)
    failed = _missing(tree, schema)
    if any(single.leaf.repeats_again for _, _, single in singles):
        raise VBAUnsupportedError("importing into a mapping of a leaf that repeats within a repeating element is not implemented")
    for sheet, (row, column), single in singles:
        text = _value_under(tree, _steps(single.leaf.path)[1:])
        value, good = (EMPTY, True) if text is None else _typed(text, single.leaf.kind)
        failed |= not good
        _put(sheet, row, column, value)
    tables: dict[int, Table] = {}
    for table, _ in columns:
        tables[id(table)] = table
    for table in tables.values():
        failed |= _fill_table(table, state, tree, overwrite and not state.append)
        if state.autofit:
            # AdjustColumnWidth: the table's columns fitted to what they hold now (tests/fixtures/xml_maps.json).
            from pyopenvba._a1 import Area

            columns_of = Range(table.sheet, [Area(table.area.top, table.area.left, table.area.bottom, table.area.right)]).vba_get("Columns")
            assert isinstance(columns_of, VBAObject)
            columns_of.vba_call("AutoFit", [])
    return 2 if failed else 0


def _missing(tree: ElementTree.Element, schema: Schema) -> bool:
    """Whether the data leaves out what the schema requires: a leaf element, or every one of the repeating element, which an
    import answers as failing validation (tests/fixtures/xml_maps.json)."""
    outer = [leaf for leaf in schema.leaves if not leaf.repeating and leaf.required and '@' not in leaf.path]
    if any(_value_under(tree, _steps(leaf.path)[1:]) is None for leaf in outer):
        return True
    if not schema.repeat_path:
        return False
    items = _items(tree, schema)
    if schema.repeat_required and not items:
        return True
    depth = len(schema.repeat_steps)
    inner = [leaf for leaf in schema.leaves if leaf.repeating and leaf.required and not leaf.repeats_again and '@' not in leaf.path]
    return any(_value_under(item, _steps(leaf.path)[depth:]) is None for item in items for leaf in inner)


def _items(tree: ElementTree.Element, schema: Schema) -> list[ElementTree.Element]:
    nodes = [tree]
    for one in schema.repeat_steps[1:]:
        nodes = [child for node in nodes for child in _children(node, one.split(":")[-1])]
    return nodes


def _fill_table(table: Table, state: MapState, tree: ElementTree.Element, overwrite: bool) -> bool:
    from pyopenvba.apps.excel._table_edits import resize
    from pyopenvba._a1 import Area

    schema = state.schema
    depth = len(schema.repeat_steps)
    items = _items(tree, schema)
    data = table.data
    if table.totals or not table.headers or data is None:
        raise VBAUnsupportedError("importing XML into a table with a totals row or no header row is not implemented")
    start = data.top if overwrite else data.bottom + 1
    rows = max(len(items), 1) if overwrite else len(items)
    bottom = start + rows - 1 if rows else data.bottom
    if bottom != table.area.bottom:
        resize(table, Area(table.area.top, table.area.left, bottom, table.area.right))
    failed = False
    for offset in range(rows):
        item = items[offset] if offset < len(items) else None
        for place, column in enumerate(table.columns):
            if column.xml is None or column.xml[0] is not state:
                continue
            leaf = column.xml[1]
            text = None if item is None else _value_under(item, _steps(leaf.path)[depth:])
            value, good = (EMPTY, True) if text is None else _typed(text, leaf.kind)
            failed |= not good
            _put(table.sheet, start + offset, table.area.left + place, value)
    return failed


def _put(sheet: Worksheet, row: int, column: int, value: object) -> None:
    from pyopenvba._a1 import Area

    Range(sheet, [Area(row, column, row, column)]).vba_set("Value", value)


# --- exporting -----------------------------------------------------------------------------------------------------------


def _written(value: object, kind: str) -> tuple[str, bool]:
    """A cell's value as an export writes it, and whether the schema takes it."""
    if value is EMPTY or value == "":
        return "", kind == "string"
    if kind == "boolean":
        if isinstance(value, bool):
            return ("true" if value else "false"), True
        return to_text(value), False
    if kind == "date":
        if isinstance(value, VBADate):
            when = value.to_datetime()
            return f"{when.year:04d}-{when.month:02d}-{when.day:02d}", True
        return to_text(value), False
    if kind in ("integer", "decimal"):
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            text = to_text(value)
            return text, kind == "decimal" or float(value) == int(float(value))
        return to_text(value), False
    return to_text(value), True


def export_text(book: Workbook, state: MapState, *, file: bool) -> tuple[str, int]:
    """What ExportXml writes into its argument, or Export and SaveAsXMLData into a file, and their answer: 1 where a value is
    not its type."""
    schema = state.schema
    if schema.namespace or any(leaf.path.rsplit("/", 1)[-1].startswith("@") for leaf in schema.leaves):
        raise VBAUnsupportedError("exporting an XML map with a namespace or attributes is not implemented")
    singles, columns = mappings(book, state)
    if not singles and not columns:
        raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
    if any(leaf.repeats_again for leaf in schema.leaves):
        raise VBAUnsupportedError("exporting an XML map with a leaf that repeats within a repeating element is not implemented")
    mapped = {single.leaf.path: (sheet, at) for sheet, at, single in singles}
    tables = {id(table): table for table, _ in columns}
    if len(tables) > 1 or (tables and singles and any(single.leaf.repeating for _, _, single in singles)):
        raise VBAUnsupportedError("exporting an XML map over several tables, or a table and its repeating leaves one by one, is not implemented")
    failed = False
    lines: list[str] = []

    def element(leaf: Leaf, value: object, indent: str) -> None:
        nonlocal failed
        text, good = _written(value, leaf.kind)
        failed |= not good
        lines.append(f"{indent}<{leaf.name}/>" if text == "" else f"{indent}<{leaf.name}>{escape(text)}</{leaf.name}>")

    def value_of(sheet: Worksheet, at: tuple[int, int]) -> object:
        from pyopenvba._a1 import Area

        return Range(sheet, [Area(at[0], at[1], at[0], at[1])]).vba_get("Value")

    repeat = schema.repeat_path
    outer = [leaf for leaf in schema.leaves if not leaf.repeating]
    inner = [leaf for leaf in schema.leaves if leaf.repeating]
    for leaf in outer:
        if leaf.path in mapped:
            element(leaf, value_of(*mapped[leaf.path]), "\t")
        elif leaf.required and "@" not in leaf.path:
            # Written without what the schema requires, which fails validation (tests/fixtures/xml_maps.json).
            failed = True
    if repeat:
        item = schema.repeat_steps[-1]
        rows: list[dict[str, object]] = []
        if tables:
            table = next(iter(tables.values()))
            if any(column.xml is None for column in table.columns) or len(table.columns) != len(inner):
                raise VBAUnsupportedError("exporting a table with columns not mapped, or leaves not mapped, is not implemented")
            data = table.data
            for row in range(data.top, data.bottom + 1) if data is not None else range(0):
                rows.append({column.xml[1].path: value_of(table.sheet, (row, table.area.left + place))
                             for place, column in enumerate(table.columns) if column.xml is not None})
        elif any(leaf.path in mapped for leaf in inner):
            rows.append({leaf.path: value_of(*mapped[leaf.path]) for leaf in inner if leaf.path in mapped})
        if rows and any(leaf.path not in rows[0] for leaf in inner):
            raise VBAUnsupportedError("exporting a repeating element some of whose leaves are not mapped is not implemented")
        for one in rows:
            lines.append(f"\t<{item}>")
            for leaf in inner:
                element(leaf, one[leaf.path], "\t\t")
            lines.append(f"\t</{item}>")
    # Lines end in CR LF, in a file and in the text ExportXml gives alike (tests/fixtures/xml_maps.json).
    end = "\r\n"
    declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + end if file else '<?xml version="1.0" standalone="yes"?>'
    body = end.join([f"<{schema.root}>", *lines, f"</{schema.root}>"])
    return declaration + body, 1 if failed else 0


# --- the objects ---------------------------------------------------------------------------------------------------------


class XmlNamespace(ExcelObject):
    vba_type_name = "XmlNamespace"

    def __init__(self, book: Workbook, uri: str, prefix: str) -> None:
        self.book, self.uri, self.prefix = book, uri, prefix

    @member(name="_Default")
    def hidden_default(self) -> object:
        """The hidden _Default, which is Value."""
        return self.Value()

    @member(default=True)
    def Value(self) -> object:
        return f'xmlns:{self.prefix}="{self.uri}"' if self.prefix else ""

    @member
    def URI(self) -> object:
        return self.uri

    @member
    def Prefix(self) -> object:
        return self.prefix

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @method
    def InstallManifest(self, *_: object) -> object:
        raise VBAUnsupportedError("XmlNamespace.InstallManifest is not implemented")


class XmlNamespaces(VBACollection, ItemDefault, ExcelObject):
    vba_type_name = "XmlNamespaces"

    def __init__(self, book: Workbook) -> None:
        refuse_unread(book)
        self.book = book

    def vba_items(self) -> list[object]:
        seen: list[object] = []
        uris: set[str] = set()
        for state in maps_of(self.book):
            if state.schema.namespace and state.schema.namespace not in uris:
                uris.add(state.schema.namespace)
                seen.append(XmlNamespace(self.book, state.schema.namespace, state.schema.prefix))
        if len(seen) > 1:
            raise VBAUnsupportedError("a workbook of XML maps in several namespaces is not implemented")
        return seen

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for one in items:
                assert isinstance(one, XmlNamespace)
                if one.prefix.casefold() == index.casefold() or one.uri == index:
                    return one
            raise error(9, "Subscript out of range")
        place = int(cast(Any, index)) if isinstance(index, (int, float)) else -1
        if not 1 <= place <= len(items):
            raise error(9, "Subscript out of range")
        return items[place - 1]

    @member
    def Value(self) -> object:
        return " ".join(cast(XmlNamespace, one).Value() for one in self.vba_items())  # type: ignore[no-untyped-call]

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @method
    def InstallManifest(self, *_: object) -> object:
        raise VBAUnsupportedError("XmlNamespaces.InstallManifest is not implemented")


class XmlSchema(ExcelObject):
    vba_type_name = "XmlSchema"

    def __init__(self, map_object: XmlMap) -> None:
        self.map_object = map_object
        self.schema = map_object.state.schema

    @member
    def Name(self) -> object:
        return self.schema.id

    @member
    def XML(self) -> object:
        return self.schema.xml

    @member
    def Namespace(self) -> object:
        return XmlNamespace(self.map_object.book, self.schema.namespace, self.schema.prefix)

    @member
    def Parent(self) -> object:
        return self.map_object

    @member
    def Application(self) -> object:
        return self.map_object.book.application


class XmlSchemas(VBACollection, ItemDefault, ExcelObject):
    vba_type_name = "XmlSchemas"

    def __init__(self, map_object: XmlMap) -> None:
        self.map_object = map_object

    def vba_items(self) -> list[object]:
        return [XmlSchema(self.map_object)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            if index.casefold() == self.map_object.state.schema.id.casefold():
                return items[0]
            raise error(9, "Subscript out of range")
        place = int(cast(Any, index)) if isinstance(index, (int, float)) else -1
        if place != 1:
            raise error(9, "Subscript out of range")
        return items[0]

    @member
    def Parent(self) -> object:
        return self.map_object

    @member
    def Application(self) -> object:
        return self.map_object.book.application


class XmlDataBinding(ExcelObject):
    vba_type_name = "XmlDataBinding"

    def __init__(self, map_object: XmlMap) -> None:
        self.map_object = map_object

    @member(name="_Default")
    def hidden_default(self) -> object:
        """The hidden _Default, which is SourceUrl."""
        return self.SourceUrl()

    @member
    def SourceUrl(self) -> object:
        return self.map_object.state.source_url

    @member
    def Parent(self) -> object:
        return self.map_object

    @member
    def Application(self) -> object:
        return self.map_object.book.application

    @method
    def Refresh(self) -> object:
        raise VBAUnsupportedError("XmlDataBinding.Refresh is not implemented")

    @method
    def LoadSettings(self, *_: object) -> object:
        raise VBAUnsupportedError("XmlDataBinding.LoadSettings is not implemented")

    @method
    def ClearSettings(self) -> object:
        raise VBAUnsupportedError("XmlDataBinding.ClearSettings is not implemented")


class XmlMap(ExcelObject):
    vba_type_name = "XmlMap"
    #: ExportXml writes its text into the caller's variable.
    vba_slot_arguments = frozenset({"exportxml"})

    def __init__(self, book: Workbook, state: MapState) -> None:
        self.book, self.state = book, state

    def vba_identity(self) -> object:
        return self.state

    def _alive(self) -> MapState:
        if not any(one is self.state for one in maps_of(self.book)):
            raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
        return self.state

    @member(name="_Default")
    def hidden_default(self) -> object:
        """The hidden _Default, which is Name."""
        return self.Name()

    @member(default=True)
    def Name(self) -> object:
        return self._alive().name

    @setter("Name")
    def _set_name(self, value: object) -> None:
        wanted = to_text(value)
        if not wanted or any(one is not self.state and one.name.casefold() == wanted.casefold() for one in maps_of(self.book)):
            raise VBAUnsupportedError("naming an XML map blank or as another is not implemented")
        self._alive().name = wanted
        self.book.xml_maps_changed = True

    @member
    def RootElementName(self) -> object:
        return self._alive().schema.root

    @member
    def RootElementNamespace(self) -> object:
        schema = self._alive().schema
        return XmlNamespace(self.book, schema.namespace, schema.prefix)

    @member
    def Schemas(self) -> object:
        return XmlSchemas(self)

    @member
    def DataBinding(self) -> object:
        return XmlDataBinding(self)

    @member
    def IsExportable(self) -> object:
        singles, columns = mappings(self.book, self._alive())
        if not singles and not columns:
            return False
        return not any(single.leaf.repeats_again for _, _, single in singles)

    def _flag(self, name: str) -> object:
        return bool(getattr(self._alive(), name))

    def _set_flag(self, name: str, value: object) -> None:
        setattr(self._alive(), name, to_bool(value))
        self.book.xml_maps_changed = True

    @member
    def ShowImportExportValidationErrors(self) -> object:
        return self._flag("show_errors")

    @setter("ShowImportExportValidationErrors")
    def _set_show(self, value: object) -> None:
        if to_bool(value):
            raise VBAUnsupportedError("XML maps that show validation errors are not implemented")
        self._set_flag("show_errors", value)

    @member
    def SaveDataSourceDefinition(self) -> object:
        return self._flag("save_source")

    @setter("SaveDataSourceDefinition")
    def _set_source(self, value: object) -> None:
        if to_bool(value):
            raise VBAUnsupportedError("XML maps that save their data source's definition are not implemented")
        self._set_flag("save_source", value)

    @member
    def AdjustColumnWidth(self) -> object:
        return self._flag("autofit")

    @setter("AdjustColumnWidth")
    def _set_autofit(self, value: object) -> None:
        self._set_flag("autofit", value)

    @member
    def PreserveColumnFilter(self) -> object:
        return self._flag("preserve_filter")

    @setter("PreserveColumnFilter")
    def _set_filter(self, value: object) -> None:
        self._set_flag("preserve_filter", value)

    @member
    def PreserveNumberFormatting(self) -> object:
        return self._flag("preserve_format")

    @setter("PreserveNumberFormatting")
    def _set_format(self, value: object) -> None:
        self._set_flag("preserve_format", value)

    @member
    def AppendOnImport(self) -> object:
        return self._flag("append")

    @setter("AppendOnImport")
    def _set_append(self, value: object) -> None:
        self._set_flag("append", value)

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application

    @member
    def WorkbookConnection(self) -> object:
        raise VBAUnsupportedError("XmlMap.WorkbookConnection is not implemented")

    @method
    def Delete(self) -> object:
        state = self._alive()
        singles, columns = mappings(self.book, state)
        if singles or columns:
            raise VBAUnsupportedError("deleting an XML map that cells are mapped to is not implemented")
        maps_of(self.book).remove(state)
        self.book.xml_maps_changed = True
        return EMPTY

    def _admit_import(self) -> None:
        if self.state.show_errors:
            raise VBAUnsupportedError("importing with validation errors shown is not implemented")

    def imported(self, data: str, overwrite: bool, url: str) -> object:
        """An import with its events: BeforeXmlImport, which can cancel it, then AfterXmlImport with the answer."""
        from pyopenvba.apps.excel import _events

        self._admit_import()
        state = self._alive()
        # Text that is not XML fails before any event is raised (tests/fixtures/xml_map_events.json).
        _parsed(data)
        if not _events.before_xml_import(self.book, self, url or "Unspecified"):
            # A handler that cancels it fails the call (tests/fixtures/xml_map_events.json).
            raise error(1004, "Application-defined or object-defined error")
        result = import_xml(self.book, state, data, overwrite)
        _events.after_xml_import(self.book, self, result)
        return VBAInt(result, "Long")

    def exported(self, url: str, *, file: bool) -> tuple[str, object] | None:
        """An export's text and answer with its events; None where BeforeXmlExport cancels it."""
        from pyopenvba.apps.excel import _events

        state = self._alive()
        if not _events.before_xml_export(self.book, self, url or "Unspecified"):
            raise error(1004, "Application-defined or object-defined error")
        text, result = export_text(self.book, state, file=file)
        _events.after_xml_export(self.book, self, url or "Unspecified", result)
        return text, VBAInt(result, "Long")

    @method
    def ImportXml(self, XmlData: object = MISSING, Overwrite: object = MISSING) -> object:  # noqa: N803 - native names
        if XmlData is MISSING:
            raise error(449, "Argument not optional")
        return self.imported(to_text(XmlData), True if Overwrite is MISSING else to_bool(Overwrite), "")

    @method
    def Import(self, Url: object = MISSING, Overwrite: object = MISSING) -> object:  # noqa: N803 - native names
        if Url is MISSING:
            raise error(449, "Argument not optional")
        text = _read_file(self.book, to_text(Url), "XmlMap.Import")
        result = self.imported(text, True if Overwrite is MISSING else to_bool(Overwrite), to_text(Url))
        self.state.source_url = to_text(Url)
        return result

    @method
    def ExportXml(self, Data: object = MISSING) -> object:  # noqa: N803 - native name
        found = self.exported("", file=False)
        if found is None:
            return VBAInt(0, "Long")
        if isinstance(Data, Slot):
            Data.set(found[0])
        return found[1]

    @method
    def Export(self, Url: object = MISSING, Overwrite: object = MISSING) -> object:  # noqa: N803 - native names
        import os

        if Url is MISSING:
            raise error(449, "Argument not optional")
        self.book.application.outside("XmlMap.Export")
        if (Overwrite is MISSING or not to_bool(Overwrite)) and os.path.exists(final_path(native_path(to_text(Url)))):
            # A file there already, and not told to overwrite it (tests/fixtures/xml_maps.json).
            raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
        found = self.exported(to_text(Url), file=True)
        if found is None:
            return VBAInt(0, "Long")
        _write_file(self.book, to_text(Url), found[0], "XmlMap.Export")
        return found[1]


class XmlMaps(VBACollection, ItemDefault, ExcelObject):
    vba_type_name = "XmlMaps"

    def __init__(self, book: Workbook) -> None:
        refuse_unread(book)
        self.book = book

    def vba_items(self) -> list[object]:
        # In the order of their names (tests/fixtures/xml_maps.json).
        return [XmlMap(self.book, state) for state in sorted(maps_of(self.book), key=lambda one: one.name.casefold())]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        if isinstance(index, str):
            for one in items:
                assert isinstance(one, XmlMap)
                # A name is found as it is spelt, case and all (tests/fixtures/xml_maps.json).
                if one.state.name == index:
                    return one
            raise error(9, "Subscript out of range")
        place = int(cast(Any, index)) if isinstance(index, (int, float)) and not isinstance(index, bool) else -1
        if not 1 <= place <= len(items):
            raise error(9, "Subscript out of range")
        return items[place - 1]

    @method
    def Add(self, Schema: object = MISSING, RootElementName: object = MISSING) -> object:  # noqa: N803 - native names
        return XmlMap(self.book, add_map(self.book, Schema, RootElementName))

    @member
    def Parent(self) -> object:
        return self.book

    @member
    def Application(self) -> object:
        return self.book.application


def plain(value: object) -> object:
    """An argument passed as the caller's variable, as the value it holds."""
    return value.get() if isinstance(value, Slot) else value


def map_state(value: object) -> MapState:
    value = plain(value)
    if not isinstance(value, XmlMap):
        raise error(13, "Type mismatch")
    return value._alive()  # pyright: ignore[reportPrivateUsage]


# --- XPath: of a cell, and of a table's column -----------------------------------------------------------------------------


class XPath(ExcelObject):
    """A cell's or a column's XML mapping. A cell inside a table's mapped column answers the column's."""

    vba_type_name = "XPath"

    def __init__(self, sheet: Worksheet, area: Area, column: tuple[Table, TableColumn] | None) -> None:
        self.sheet, self.area, self.column = sheet, area, column

    def _found(self) -> tuple[MapState, Leaf, bool] | None:
        """The mapping: its map, leaf, and whether it repeats (a table's column)."""
        from pyopenvba.apps.excel._tables import table_at

        column: tuple[Table, TableColumn] | None = self.column
        if column is None and self.area.rows == 1:
            table = table_at(self.sheet, self.area.top, self.area.left)
            if table is not None:
                column = (table, table.columns[self.area.left - table.area.left])
        if column is not None:
            mapped = cast("tuple[MapState, Leaf] | None", column[1].xml)
            return None if mapped is None else (mapped[0], mapped[1], True)
        single = self.sheet.xml_cells.get((self.area.top, self.area.left))
        return None if single is None else (single.map, single.leaf, False)

    @member(name="_Default")
    def hidden_default(self) -> object:
        """The hidden _Default, which is Value."""
        return self.Value()

    @member(default=True)
    def Value(self) -> object:
        found = self._found()
        return "" if found is None else found[1].path

    @member
    def Map(self) -> object:
        found = self._found()
        if found is None:
            if self.column is not None:
                raise error(_UNSPECIFIED, "Automation error\nUnspecified error")
            raise error(_UNSPECIFIED, "Automation error")
        return XmlMap(self.sheet.book, found[0])

    @member
    def Repeating(self) -> object:
        found = self._found()
        if found is None:
            if self.column is not None:
                raise error(1004, "Unable to get the Repeating property of the XPath class")
            raise error(_UNSPECIFIED, "Automation error")
        return found[2]

    @member
    def Parent(self) -> object:
        return Range(self.sheet, [self.area])

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @method
    def SetValue(self, Map: object = MISSING, XPath: object = MISSING, SelectionNamespace: object = MISSING,  # noqa: N803
                 Repeating: object = MISSING) -> object:
        from pyopenvba.apps.excel._tables import guid, table_at

        if Map is MISSING or XPath is MISSING:
            raise error(449, "Argument not optional")
        if SelectionNamespace is not MISSING or Repeating is not MISSING:
            raise VBAUnsupportedError("XPath.SetValue with SelectionNamespace or Repeating is not implemented")
        state = map_state(Map)
        leaf = _leaf_at(state.schema, to_text(XPath))
        book = self.sheet.book
        if leaf.kind not in _TYPES:
            raise VBAUnsupportedError(f"mapping an XML leaf of type {leaf.kind} is not implemented")
        if self.column is not None:
            self._map_column(self.column[0], self.column[1], state, leaf)
        elif self.area.rows == 1 and self.area.columns == 1:
            if table_at(self.sheet, self.area.top, self.area.left) is not None:
                raise VBAUnsupportedError("mapping a cell of a table through Range.XPath is not implemented")
            self.sheet.xml_cells[(self.area.top, self.area.left)] = SingleMapping(state, leaf, guid(), guid())
        elif self.area.columns == 1 and leaf.repeating and not leaf.repeats_again:
            self._block_table(state, leaf)
        else:
            raise VBAUnsupportedError("mapping a block of cells to an XML leaf that does not repeat is not implemented")
        book.xml_maps_changed = True
        self.sheet.touched()
        return EMPTY

    def _map_column(self, table: Table, column: TableColumn, state: MapState, leaf: Leaf) -> None:
        if table.part:
            raise VBAUnsupportedError("mapping a column of a table read from a file is not implemented")
        if not leaf.repeating or leaf.repeats_again:
            raise VBAUnsupportedError("mapping a table's column to an XML leaf that does not repeat once is not implemented")
        others = [one.xml for one in table.columns if one.xml is not None and one is not column]
        if any(other[0] is not state for other in others):
            raise VBAUnsupportedError("mapping a table's columns to several XML maps is not implemented")
        if table.kind not in ("", "xml"):
            raise VBAUnsupportedError(f"mapping a column of a table of type {table.kind!r} is not implemented")
        column.xml, column.xml_name = (state, leaf), leaf.name
        table.kind = "xml"
        table.changed = True

    def _block_table(self, state: MapState, leaf: Leaf) -> None:
        """A block of one column mapped at once: a table of it, its header the leaf's name (tests/fixtures/xml_maps.json)."""
        from pyopenvba.apps.excel._tables import add

        _put(self.sheet, self.area.top, self.area.left, leaf.name)
        table = add(self.sheet, Range(self.sheet, [self.area]), 1, "TableStyleMedium2")
        self._map_column(table, table.columns[0], state, leaf)

    @method
    def Clear(self) -> object:
        from pyopenvba.apps.excel._tables import table_at

        column: tuple[Table, TableColumn] | None = self.column
        if column is None and self.area.rows == 1:
            table = table_at(self.sheet, self.area.top, self.area.left)
            if table is not None:
                column = (table, table.columns[self.area.left - table.area.left])
        if column is not None:
            if column[1].xml is not None:
                if column[0].part:
                    raise VBAUnsupportedError("clearing a mapping of a table read from a file is not implemented")
                column[1].xml = None
                column[0].changed = True
                self.sheet.book.xml_maps_changed = True
            return EMPTY
        if self.sheet.xml_cells.pop((self.area.top, self.area.left), None) is not None:
            self.sheet.book.xml_maps_changed = True
            self.sheet.touched()
        return EMPTY


def cell_xpath(target: Range) -> XPath:
    """Range.XPath: of one cell or of one column of cells; a wider block is out of range (tests/fixtures/unassessed_more.json)."""
    refuse_unread(target.sheet.book)
    if len(target.areas) != 1 or target.first.columns != 1:
        raise error(-2147024809, "The specified value is out of range.")
    return XPath(target.sheet, target.first, None)


def column_xpath(table: Table, column: TableColumn) -> XPath:
    from pyopenvba._a1 import Area

    refuse_unread(table.sheet.book)
    return XPath(table.sheet, Area(table.area.top, table.area.left, table.area.top, table.area.left), (table, column))


def table_map(table: Table) -> object:
    """ListObject.XmlMap: the map its columns are mapped to, or Nothing."""
    refuse_unread(table.sheet.book)
    found = next((column.xml[0] for column in table.columns if column.xml is not None), None)
    return NOTHING if found is None else XmlMap(table.sheet.book, found)


# --- the workbook's members --------------------------------------------------------------------------------------------


def workbook_import(book: Workbook, data: str, import_map: object, overwrite: object, destination: object, *, url: str = "") -> object:
    """Workbook.XmlImportXml and XmlImport with a map; with none, Excel infers a schema, which is not built."""
    held = import_map.get() if isinstance(import_map, Slot) else import_map
    if not isinstance(held, XmlMap):
        raise VBAUnsupportedError("importing XML with no map, which Excel infers a schema for, is not implemented")
    if destination is not MISSING:
        raise VBAUnsupportedError("importing XML into a map with a Destination is not implemented")
    if overwrite is MISSING:
        raise VBAUnsupportedError("importing XML into a map without saying Overwrite is not implemented")
    state = map_state(held)
    result = held.imported(data, to_bool(overwrite), url)
    if url:
        state.source_url = url
    return result


def read_url(book: Workbook, url: object, what: str) -> tuple[str, str]:
    if url is MISSING:
        raise error(449, "Argument not optional")
    return _read_file(book, to_text(url), what), to_text(url)


def save_as_xml_data(book: Workbook, filename: object, map_object: object) -> None:
    if filename is MISSING or map_object is MISSING:
        raise error(449, "Argument not optional")
    book.application.outside("Workbook.SaveAsXMLData")
    map_state(map_object)
    held = plain(map_object)
    assert isinstance(held, XmlMap)
    found = held.exported(to_text(filename), file=True)
    if found is not None:
        _write_file(book, to_text(filename), found[0], "Workbook.SaveAsXMLData")


# --- the file -------------------------------------------------------------------------------------------------------------

_SINGLES_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml"
_REL_MAPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps"
_REL_SINGLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
_SINGLES_NAMESPACES = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                       'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr xr3 xr6" '
                       'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
                       'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
                       'xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6"')


def _schema_written(schema: Schema) -> str:
    """The schema as xmlMaps.xml holds it: as given, with xmlns="" on its root where it names no default namespace."""
    text = schema.xml.strip()
    found = re.match(r"<[^>]*?(/?)>", text)
    if found is None or re.search(r'\sxmlns="', found.group(0)):
        return text
    end = found.end() - (2 if found.group(1) else 1)
    return text[:end] + ' xmlns=""' + text[end:]


def maps_xml(book: Workbook) -> str:
    namespaces = " ".join(f"xmlns:{state.schema.prefix}='{state.schema.namespace}'" for state in maps_of(book) if state.schema.namespace)
    schemas = "".join(f'<Schema ID="{state.schema.id}"' + (f' Namespace="{escape(state.schema.namespace)}"' if state.schema.namespace else "")
                      + f">{_schema_written(state.schema)}</Schema>" for state in sorted(maps_of(book), key=lambda one: one.schema.id))
    maps = "".join(f'<Map ID="{state.id}" Name="{escape(state.name)}" RootElement="{escape(state.schema.root)}" SchemaID="{state.schema.id}" '
                   f'ShowImportExportValidationErrors="{str(state.show_errors).lower()}" AutoFit="{str(state.autofit).lower()}" '
                   f'Append="{str(state.append).lower()}" PreserveSortAFLayout="{str(state.preserve_filter).lower()}" '
                   f'PreserveFormat="{str(state.preserve_format).lower()}"/>'
                   for state in sorted(maps_of(book), key=lambda one: one.name.casefold()))
    return (f'<MapInfo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" SelectionNamespaces="{escape(namespaces)}">'
            f"{schemas}{maps}</MapInfo>\r\n")


def singles_xml(sheet: Worksheet) -> str:
    cells: list[str] = []
    for number, ((row, column), single) in enumerate(sorted(sheet.xml_cells.items(), key=lambda item: (item[0][0], item[0][1])), 1):
        from pyopenvba._a1 import Area

        where = Area(row, column, row, column).address(absolute=False)
        cells.append(f'<singleXmlCell id="{number}" xr6:uid="{single.uid}" r="{where}" connectionId="0"><xmlCellPr id="1" '
                     f'xr6:uid="{single.cell_uid}" uniqueName="1"><xmlPr mapId="{single.map.id}" xpath="{escape(single.leaf.path)}" '
                     f'xmlDataType="{single.leaf.kind}"/></xmlCellPr></singleXmlCell>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<singleXmlCells {_SINGLES_NAMESPACES}>{"".join(cells)}</singleXmlCells>'


def write_parts(book: Workbook, package: Any, add_sheet_relationship: Any) -> None:
    """xmlMaps.xml and each sheet's single mappings, where the macro changed them."""
    from pyopenvba.powerquery._sheets import add_content_type, add_relationship

    if not book.xml_maps_changed:
        return
    if any(state.source_url for state in maps_of(book)):
        # Excel keeps a map imported from a file as a workbook connection to it, and the map bound to that connection.
        raise VBAUnsupportedError("saving an XML map imported from a file, which Excel keeps as a workbook connection, is not implemented")
    if maps_of(book):
        if not package.has("xl/xmlMaps.xml"):
            add_relationship(package, "xl/_rels/workbook.xml.rels", _REL_MAPS, "xmlMaps.xml")
        package.write("xl/xmlMaps.xml", maps_xml(book).encode("utf-8"))
    elif package.has("xl/xmlMaps.xml"):
        raise VBAUnsupportedError("saving a workbook whose XML maps were all deleted is not implemented")
    for sheet in book.sheets_:
        if not sheet.xml_cells and not sheet.xml_cells_part:
            continue
        if not sheet.xml_cells:
            raise VBAUnsupportedError("saving a sheet whose single XML mappings were all cleared is not implemented")
        if not sheet.xml_cells_part:
            number = 1
            while package.has(f"xl/tables/tableSingleCells{number}.xml"):
                number += 1
            sheet.xml_cells_part = f"xl/tables/tableSingleCells{number}.xml"
            add_content_type(package, sheet.xml_cells_part, _SINGLES_TYPE)
            add_sheet_relationship(package, sheet.part_name, f"../tables/tableSingleCells{number}.xml", _REL_SINGLES)
        package.write(sheet.xml_cells_part, singles_xml(sheet).encode("utf-8"))
    book.xml_maps_changed = False


def read_parts(book: Workbook, package: Any, sheet_parts: dict[str, Worksheet]) -> None:
    """The maps, and each sheet's single mappings and mapped columns, from a file Excel saved. Maps the model does not
    build are not read at all: the file opens, its parts stay as they came, and every XML map member says why."""
    if not package.has("xl/xmlMaps.xml"):
        return
    try:
        _read_parts(book, package)
    except (VBAUnsupportedError, KeyError, ValueError, ElementTree.ParseError) as failed:
        book.xml_maps.clear()
        for sheet in book.sheets_:
            sheet.xml_cells.clear()
            for table in sheet.tables:
                for column in table.columns:
                    column.xml, column.xml_name = None, ""
        book.xml_maps_unread = str(failed) or type(failed).__name__
    book.xml_maps_changed = False


def refuse_unread(book: Workbook) -> None:
    """Report maps a file has that the model left unread (see read_parts)."""
    if book.xml_maps_unread:
        raise VBAUnsupportedError(f"the XML maps of this file are not implemented: {book.xml_maps_unread}")


def _read_parts(book: Workbook, package: Any) -> None:
    root = ElementTree.fromstring(package.read("xl/xmlMaps.xml"))
    main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    schemas: dict[str, Schema] = {}
    raw = package.read("xl/xmlMaps.xml").decode("utf-8")
    for found in re.finditer(r'<Schema ID="([^"]+)"[^>]*>(.*?)</Schema>', raw, re.DOTALL):
        text = re.sub(r'(<[^>]*?)\sxmlns=""', r"\1", found.group(2), count=1)
        number = int(re.sub(r"\D", "", found.group(1)) or 0)
        schemas[found.group(1)] = read_schema(text, MISSING, number)
    for element in root.findall(main + "Map"):
        schema = schemas.get(element.get("SchemaID", ""))
        if schema is None:
            raise VBAUnsupportedError("an XML map whose schema the file does not hold is not implemented")
        if element.get("ShowImportExportValidationErrors") == "true":
            raise VBAUnsupportedError("an XML map that shows validation errors")
        if element.find(main + "DataBinding") is not None:
            raise VBAUnsupportedError("an XML map bound to a data source")
        if schema.root != element.get("RootElement"):
            schema = read_schema(schema.xml, element.get("RootElement", ""), int(re.sub(r"\D", "", schema.id) or 0))
        state = MapState(id=int(element.get("ID", "0")), name=element.get("Name", ""), schema=schema,
                         show_errors=element.get("ShowImportExportValidationErrors") == "true", autofit=element.get("AutoFit") != "false",
                         append=element.get("Append") == "true", preserve_filter=element.get("PreserveSortAFLayout") != "false",
                         preserve_format=element.get("PreserveFormat") != "false")
        maps_of(book).append(state)
    maps_of(book).sort(key=lambda one: one.id)
    book.xml_maps_made = max((one.id for one in maps_of(book)), default=0)
    book.xml_schemas_made = max((int(re.sub(r"\D", "", one.schema.id) or 0) for one in maps_of(book)), default=0)
    by_id = {one.id: one for one in maps_of(book)}
    for sheet in book.sheets_:
        for table in sheet.tables:
            for column, found in zip(table.columns, re.finditer(r"<tableColumn\b.*?(?:/>|</tableColumn>)", table.xml, re.DOTALL)):
                mapped = re.search(r'<xmlColumnPr mapId="(\d+)" xpath="([^"]*)"', found.group(0))
                if mapped is not None:
                    state = by_id[int(mapped.group(1))]
                    column.xml = (state, _leaf_at(state.schema, mapped.group(2)))
                    column.xml_name = column.xml[1].name
        part = sheet.xml_cells_part
        if part and package.has(part):
            from pyopenvba._a1 import parse_area

            text = package.read(part).decode("utf-8")
            for cell in re.finditer(r'<singleXmlCell\b[^>]*\bxr6:uid="([^"]*)"[^>]*\br="([^"]+)".*?<xmlCellPr\b[^>]*\bxr6:uid="([^"]*)".*?'
                                    r'<xmlPr mapId="(\d+)" xpath="([^"]*)"', text, re.DOTALL):
                state = by_id[int(cell.group(4))]
                at = parse_area(cell.group(2))
                sheet.xml_cells[(at.top, at.left)] = SingleMapping(state, _leaf_at(state.schema, cell.group(5)), cell.group(1), cell.group(3))
    book.xml_maps_changed = False


__all__ = ["XPath", "XmlMap", "XmlMaps", "XmlNamespaces", "cell_xpath", "column_xpath", "read_parts", "table_map", "write_parts"]
