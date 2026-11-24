"""A custom XML part's nodes and what goes with them -- CustomXMLNode, CustomXMLNodes, the NamespaceManager's
CustomXMLPrefixMappings, Errors and SchemaCollection -- and the parts a save writes under customXml/, as live Excel
has them (scripts/measure_custom_xml.py, tests/fixtures/custom_xml*.json).

- A part with a default namespace holds the prefix ns0 for it in its NamespaceManager, then xml; AddNamespace adds
  to the end, and a prefix looked up finds its last mapping, a namespace its first.
- A node's XPath names each step with its place among its like: /ns0:root[1]/ns0:x[2], /book[1]/title[1]/@lang,
  .../text()[1], .../comment()[1], .../processing-instruction('pi')[1].
- An element's NodeValue is its XML; an attribute's, text's, a comment's and an instruction's is its value. An
  attribute has child nodes by HasChildNodes but no ChildNodes.
- Node edits a macro makes raise none of CustomXMLPart's node events.
- A save writes each part a macro added, loaded or changed as customXml/itemN.xml, with itemPropsN.xml naming the
  part's ID and its namespace, the item's relationship to it, the workbook's relationship to the item and the
  properties' content type.
"""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._custom_xml_tree import (ATTRIBUTE, COMMENT, DOCUMENT, ELEMENT, INSTRUCTION, TEXT,
                                                  XML_NAMESPACE, Node, parsed, serialized)
from pyopenvba.apps.excel._custom_xml_xpath import selected, undeclared
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import MISSING, NOTHING, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.powerquery._opc import OpcFile
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.apps.excel._workbook_more import CustomXMLPart

INVALID = (-536602618, 'This operation cannot be performed because it would result in structurally invalid XML.')
EVENTS_FAILED = 'This operation failed because one or more changes made during events were invalid.'
PROPERTIES_TYPE = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
ITEM_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
PROPERTIES_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
_DATASTORE = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'


class _OfficeObject(ExcelObject):
    """One of Office's custom XML classes, which the member inventory leaves out: a member its type library does not
    list is error 438, as Excel answers it; one it lists and the model lacks reports itself."""

    vba_library = 'office'

    def vba_no_member(self, name: str) -> Exception:
        return no_member(self, name) or super().vba_no_member(name)


def no_member(target: ExcelObject, name: str) -> Exception | None:
    """Error 438 with Excel's words for a member Office's type library does not list on the object's class."""
    from pyopenvba.interpreter._type_library import find_class

    found = find_class('office', target.vba_type_name) or find_class('office', '_' + target.vba_type_name)
    if found is not None and name.lower() not in found.members:
        return error(438, "Object doesn't support this property or method")
    return None


def view(part: CustomXMLPart, node: Node | None) -> object:
    """The CustomXMLNode a macro holds for ``node``, the same each time; Nothing for none."""
    if node is None:
        return NOTHING
    if node.view is None:
        node.view = CustomXMLNode(part, node)
    return node.view


def mappings_for(document: Node | None) -> list[tuple[str, str]]:
    """The prefixes a part's NamespaceManager starts with: ns0 for its element's default namespace, then xml. A
    default namespace an element inside declares gets none here, as was not measured (see _qualified)."""
    found: list[tuple[str, str]] = []
    root = document.root() if document is not None else None
    if root is not None:
        for node in [root, *_elements(root)]:
            for attribute in node.attributes:
                if id(attribute) in node.declarations and attribute.name != 'xmlns':
                    raise VBAUnsupportedError('custom XML declaring a namespace prefix is not implemented')
        if root.uri:
            found.append(('ns0', root.uri))
    found.append(('xml', XML_NAMESPACE))
    return found


def _elements(node: Node) -> list[Node]:
    found: list[Node] = []
    for child in node.children:
        if child.kind == ELEMENT:
            found += [child, *_elements(child)]
    return found


def resolver(part: CustomXMLPart):  # type: ignore[no-untyped-def]
    def resolve(prefix: str) -> str:
        for name, uri in reversed(part.mappings):
            if name == prefix:
                return uri
        raise undeclared(prefix)
    return resolve


def select(part: CustomXMLPart, context: Node, path: object) -> list[Node]:
    return selected(to_text(path), context, resolver(part))


class CustomXMLNodes(VBACollection, _OfficeObject):
    vba_type_name = 'CustomXMLNodes'

    def __init__(self, part: CustomXMLPart, nodes: list[Node]) -> None:
        self.part, self.nodes = part, nodes

    def vba_items(self) -> list[object]:
        return [view(self.part, node) for node in self.nodes]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(9, 'Subscript out of range')
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.part

    @member
    def Application(self) -> object:
        return self.part.book.application


class CustomXMLNode(_OfficeObject):
    vba_type_name = 'CustomXMLNode'

    def __init__(self, part: CustomXMLPart, node: Node) -> None:
        self.part, self.node = part, node

    @member
    def Parent(self) -> object:
        return self.part

    @member
    def Application(self) -> object:
        return self.part.book.application

    @member
    def BaseName(self) -> object:
        node = self.node
        return {TEXT: '#text', COMMENT: '#comment', DOCUMENT: '#document'}.get(node.kind, node.local)

    @member
    def NamespaceURI(self) -> object:
        return self.node.uri

    @member
    def NodeType(self) -> object:
        return VBAInt(self.node.kind, 'Long')

    @member
    def NodeValue(self) -> object:
        if self.node.kind in (ELEMENT, DOCUMENT):
            return serialized(self.node)
        return self.node.value

    @setter('NodeValue')
    def _set_node_value(self, value: object) -> None:
        if self.node.kind in (ELEMENT, DOCUMENT):
            raise VBAUnsupportedError("setting an element's NodeValue is not implemented")
        self._set_value(to_text(value).replace('\r\n', '\n'))

    def _set_value(self, text: str) -> None:
        """An attribute given a value goes last among its element's, as Excel puts it (tests/fixtures/custom_xml1.json)."""
        node = self.node
        node.value = text
        if node.kind == ATTRIBUTE and node.parent is not None:
            node.parent.attributes.remove(node)
            node.parent.attributes.append(node)
        self.part.edited()

    @member
    def Text(self) -> object:
        return self.node.text()

    @setter('Text')
    def _set_text(self, value: object) -> None:
        text = to_text(value).replace('\r\n', '\n')
        if self.node.kind == ELEMENT:
            self.node.children = []
            if text:
                self.node.append(Node(TEXT, '#text', text))
        elif self.node.kind == DOCUMENT:
            raise VBAUnsupportedError("setting a document node's Text is not implemented")
        else:
            self._set_value(text)
            return
        self.part.edited()

    @member
    def XML(self) -> object:
        return serialized(self.node)

    @member
    def XPath(self) -> object:
        return xpath_of(self.part, self.node)

    @member
    def Attributes(self) -> object:
        # A node that is not an element has none, but a collection all the same (tests/fixtures/custom_xml1.json).
        return CustomXMLNodes(self.part, self.node.plain_attributes() if self.node.kind == ELEMENT else [])

    @member
    def ChildNodes(self) -> object:
        if self.node.kind == ATTRIBUTE:
            return NOTHING
        return CustomXMLNodes(self.part, list(self.node.children))

    @member
    def FirstChild(self) -> object:
        children = self.node.children if self.node.kind != ATTRIBUTE else []
        return view(self.part, children[0] if children else None)

    @member
    def LastChild(self) -> object:
        children = self.node.children if self.node.kind != ATTRIBUTE else []
        return view(self.part, children[-1] if children else None)

    @member
    def HasChildNodes(self) -> object:
        # An attribute holds its value as a child, which ChildNodes does not hand over.
        return self.node.kind == ATTRIBUTE or bool(self.node.children)

    def _sibling(self, step: int) -> object:
        parent = self.node.parent
        if parent is None or self.node.kind == ATTRIBUTE:
            return NOTHING
        place = parent.children.index(self.node) + step
        return view(self.part, parent.children[place] if 0 <= place < len(parent.children) else None)

    @member
    def NextSibling(self) -> object:
        return self._sibling(1)

    @member
    def PreviousSibling(self) -> object:
        return self._sibling(-1)

    @member
    def ParentNode(self) -> object:
        return view(self.part, self.node.parent)

    @member
    def OwnerPart(self) -> object:
        return self.part

    @member
    def OwnerDocument(self) -> object:
        return self.part.book

    @method
    def SelectNodes(self, XPath: object = MISSING) -> object:
        return CustomXMLNodes(self.part, select(self.part, self.node, XPath))

    @method
    def SelectSingleNode(self, XPath: object = MISSING) -> object:
        found = select(self.part, self.node, XPath)
        return view(self.part, found[0] if found else None)

    @method
    def Delete(self) -> object:
        node = self.node
        parent = node.parent
        if parent is None or parent.kind == DOCUMENT and node.kind == ELEMENT:
            raise error(-2147024809, EVENTS_FAILED)
        if node.kind == ATTRIBUTE:
            parent.attributes.remove(node)
        else:
            parent.children.remove(node)
        node.parent = None
        self.part.edited()
        return None

    @method
    def AppendChildNode(self, Name: object = MISSING, NamespaceURI: object = MISSING, NodeType: object = MISSING,
                        NodeValue: object = MISSING) -> object:
        add(self.part, self.node, Name, NamespaceURI, None, NodeType, NodeValue)
        return None

    @method
    def AppendChildSubtree(self, XML: object = MISSING) -> object:
        _graft(self.part, self.node, to_text(XML), None)
        return None

    @method
    def InsertNodeBefore(self, Name: object = MISSING, NamespaceURI: object = MISSING, NodeType: object = MISSING,
                         NodeValue: object = MISSING, NextSibling: object = MISSING) -> object:
        add(self.part, self.node, Name, NamespaceURI, _child(self.node, NextSibling), NodeType, NodeValue)
        return None

    @method
    def InsertSubtreeBefore(self, XML: object = MISSING, NextSibling: object = MISSING) -> object:
        _graft(self.part, self.node, to_text(XML), _child(self.node, NextSibling))
        return None

    @method
    def RemoveChild(self, Child: object = MISSING) -> object:
        child = _child(self.node, Child)
        if child is None:
            raise error(449, 'Argument not optional')
        self.node.children.remove(child)
        child.parent = None
        self.part.edited()
        return None

    @method
    def ReplaceChildNode(self, OldNode: object = MISSING, Name: object = MISSING, NamespaceURI: object = MISSING,
                         NodeType: object = MISSING, NodeValue: object = MISSING) -> object:
        old = _child(self.node, OldNode)
        if old is None:
            raise error(449, 'Argument not optional')
        add(self.part, self.node, Name, NamespaceURI, old, NodeType, NodeValue)
        self.node.children.remove(old)
        old.parent = None
        self.part.edited()
        return None

    @method
    def ReplaceChildSubtree(self, XML: object = MISSING, OldNode: object = MISSING) -> object:
        old = _child(self.node, OldNode)
        if old is None:
            raise error(449, 'Argument not optional')
        _graft(self.part, self.node, to_text(XML), old)
        self.node.children.remove(old)
        old.parent = None
        self.part.edited()
        return None


def _child(parent: Node, given: object) -> Node | None:
    """The node a macro passed, which has to be one of ``parent``'s children; None for none."""
    if given is MISSING or given is NOTHING:
        return None
    if not isinstance(given, CustomXMLNode):
        raise error(13, 'Type mismatch')
    if given.node.parent is not parent or given.node.kind == ATTRIBUTE:
        raise VBAUnsupportedError('a node that is not one of the parent\'s children is not implemented')
    return given.node


def add(part: CustomXMLPart, parent: Node, name: object, uri: object, before: Node | None, kind: object,
        value: object) -> Node:
    """AddNode, AppendChildNode, InsertNodeBefore and ReplaceChildNode's new node."""
    wanted = ELEMENT if kind is MISSING else int(to_integer(kind, 'Long'))
    name_text = '' if name is MISSING else to_text(name)
    uri_text = '' if uri is MISSING else to_text(uri)
    text = '' if value is MISSING else to_text(value).replace('\r\n', '\n')
    if parent.kind not in (ELEMENT, DOCUMENT):
        raise error(*INVALID)
    if wanted == ELEMENT:
        if not name_text or parent.kind == DOCUMENT and parent.root() is not None:
            raise error(*INVALID)
        made = Node(ELEMENT, name_text, uri=uri_text)
        made.parent = parent
        if uri_text != parent.namespace_of(''):
            declaration = Node(ATTRIBUTE, 'xmlns', uri_text)
            declaration.parent = made
            made.attributes.append(declaration)
            made.declarations.add(id(declaration))
        if text:
            made.append(Node(TEXT, '#text', text))
    elif wanted == ATTRIBUTE:
        if not name_text or parent.kind != ELEMENT or uri_text or ':' in name_text:
            raise VBAUnsupportedError('an attribute with no name, in a namespace or on no element is not implemented')
        if any(one.name == name_text for one in parent.plain_attributes()):
            raise VBAUnsupportedError('an attribute an element has already is not implemented')
        made = Node(ATTRIBUTE, name_text, text)
        made.parent = parent
        parent.attributes.append(made)
        part.edited()
        return made
    elif wanted in (TEXT, COMMENT):
        made = Node(wanted, '#text' if wanted == TEXT else '#comment', text)
    elif wanted == INSTRUCTION:
        if not name_text:
            raise error(*INVALID)
        made = Node(INSTRUCTION, name_text, text)
    else:
        raise VBAUnsupportedError(f'a custom XML node of type {wanted} is not implemented')
    parent.append(made, before)
    part.edited()
    return made


def _graft(part: CustomXMLPart, parent: Node, text: str, before: Node | None) -> None:
    """A subtree of XML put into ``parent``."""
    if parent.kind != ELEMENT:
        raise VBAUnsupportedError('a subtree put anywhere but into an element is not implemented')
    document = parsed(text)
    if document.declared or any(one.kind != ELEMENT for one in document.children):
        raise VBAUnsupportedError('a subtree with more than an element in it is not implemented')
    for node in document.children:
        node.grafted = True
        parent.append(node, before)
    if mappings_for(document)[:-1]:
        raise VBAUnsupportedError('a subtree in a namespace is not implemented')
    part.edited()


def xpath_of(part: CustomXMLPart, node: Node) -> str:
    """The XPath a node reads: each step from the document down with its place among its like."""
    if node.kind == DOCUMENT or node.parent is None:
        return ''
    above = xpath_of(part, node.parent)
    if node.kind == ATTRIBUTE:
        return f'{above}/@{_qualified(part, node)}'
    siblings = node.parent.children
    if node.kind == ELEMENT:
        like = [one for one in siblings if one.kind == ELEMENT and one.local == node.local and one.uri == node.uri]
        step = _qualified(part, node)
    elif node.kind == TEXT:
        like, step = [one for one in siblings if one.kind == TEXT], 'text()'
    elif node.kind == COMMENT:
        like, step = [one for one in siblings if one.kind == COMMENT], 'comment()'
    else:
        like = [one for one in siblings if one.kind == INSTRUCTION and one.name == node.name]
        step = f"processing-instruction('{node.name}')"
    return f'{above}/{step}[{like.index(node) + 1}]'


def _qualified(part: CustomXMLPart, node: Node) -> str:
    if not node.uri:
        return node.local
    for prefix, uri in part.mappings:
        if uri == node.uri:
            return f'{prefix}:{node.local}'
    raise VBAUnsupportedError('the XPath of a node in a namespace the part maps no prefix to is not implemented')


class CustomXMLPrefixMapping(_OfficeObject):
    vba_type_name = 'CustomXMLPrefixMapping'

    def __init__(self, owner: CustomXMLPrefixMappings, prefix: str, uri: str) -> None:
        self.owner, self.prefix, self.uri = owner, prefix, uri

    @member
    def Prefix(self) -> object:
        return self.prefix

    @member
    def NamespaceURI(self) -> object:
        return self.uri

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.part.book.application


class CustomXMLPrefixMappings(VBACollection, _OfficeObject):
    vba_type_name = 'CustomXMLPrefixMappings'

    def __init__(self, part: CustomXMLPart) -> None:
        self.part = part

    def vba_items(self) -> list[object]:
        return [CustomXMLPrefixMapping(self, prefix, uri) for prefix, uri in self.part.mappings]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(9, 'Subscript out of range')
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.part

    @member
    def Application(self) -> object:
        return self.part.book.application

    @method
    def AddNamespace(self, Prefix: object = MISSING, NamespaceURI: object = MISSING) -> object:
        """A prefix added last; one already there takes the new namespace where it stands
        (tests/fixtures/custom_xml_b.json)."""
        prefix, uri = to_text(Prefix), to_text(NamespaceURI)
        mappings = self.part.mappings
        place = next((at for at, (name, _) in enumerate(mappings) if name == prefix), None)
        if place is None:
            mappings.append((prefix, uri))
        else:
            mappings[place] = (prefix, uri)
        return None

    @method
    def LookupNamespace(self, Prefix: object = MISSING) -> object:
        wanted = to_text(Prefix)
        return next((uri for prefix, uri in reversed(self.part.mappings) if prefix == wanted), '')

    @method
    def LookupPrefix(self, NamespaceURI: object = MISSING) -> object:
        wanted = to_text(NamespaceURI)
        return next((prefix for prefix, uri in self.part.mappings if uri == wanted), '')


class CustomXMLValidationErrors(VBACollection, _OfficeObject):
    """A part's validation errors: none, as no schema a macro gives is checked."""

    vba_type_name = 'CustomXMLValidationErrors'

    def __init__(self, part: CustomXMLPart) -> None:
        self.part = part

    def vba_items(self) -> list[object]:
        return []

    def vba_lookup(self, index: object, items: list[object]) -> object:
        raise error(9, 'Subscript out of range')

    @member
    def Parent(self) -> object:
        return self.part

    @member
    def Application(self) -> object:
        return self.part.book.application


class CustomXMLSchema(_OfficeObject):
    vba_type_name = 'CustomXMLSchema'

    def __init__(self, owner: CustomXMLSchemaCollection, uri: str) -> None:
        self.owner, self.uri = owner, uri

    @member
    def NamespaceURI(self) -> object:
        return self.uri

    @member
    def Location(self) -> object:
        return ''

    @member
    def Parent(self) -> object:
        return self.owner

    @member
    def Application(self) -> object:
        return self.owner.part.book.application


class CustomXMLSchemaCollection(VBACollection, _OfficeObject):
    """A part's schemas: one for its namespace, with no location, where it has one."""

    vba_type_name = 'CustomXMLSchemaCollection'

    def __init__(self, part: CustomXMLPart) -> None:
        self.part = part

    def vba_items(self) -> list[object]:
        return [CustomXMLSchema(self, uri) for uri in self.part.schemas]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        place = int(to_integer(index, 'Long'))
        if not 1 <= place <= len(items):
            raise error(-2147352565, EVENTS_FAILED)
        return items[place - 1]

    @member
    def Parent(self) -> object:
        return self.part

    @member
    def Application(self) -> object:
        return self.part.book.application

    @method
    def Validate(self) -> object:
        return True

    @method
    def Add(self, NamespaceURI: object = MISSING, Alias: object = MISSING, FileName: object = MISSING,
            InstallForAllUsers: object = MISSING) -> object:
        raise VBAUnsupportedError('adding a schema to a custom XML part is not implemented')


# --- saving ------------------------------------------------------------------------------------------------------


def namespaces(document: Node | None) -> list[str]:
    """The namespaces a document's elements are in when it is loaded, in the order they first come: the part's
    schemas, and the schema references its properties name. Elements added later add none
    (tests/fixtures/custom_xml_b.json, _d)."""
    root = document.root() if document is not None else None
    found: list[str] = []
    for node in [root, *_elements(root)] if root is not None else []:
        if node.uri and node.uri not in found:
            found.append(node.uri)
    return found


def properties_part(identity: str, uris: list[str], empty: bool) -> str:
    """itemPropsN.xml: the part's ID, and each of its namespaces as a schema reference."""
    head = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n'
    store = f'<ds:datastoreItem ds:itemID="{identity}" xmlns:ds="{_DATASTORE}"'
    if empty:
        return head + store + '/>'
    refs = ('<ds:schemaRefs>' + ''.join(f'<ds:schemaRef ds:uri="{uri}"/>' for uri in uris) + '</ds:schemaRefs>'
            if uris else '<ds:schemaRefs/>')
    return head + store + '>' + refs + '</ds:datastoreItem>'


def item_relationships(number: int) -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.'
            'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'
            f'{PROPERTIES_RELATIONSHIP}" Target="itemProps{number}.xml"/></Relationships>')


def write(book: Workbook, package: OpcFile) -> None:
    """Every part a macro added, loaded, changed or deleted written again under customXml/, as Excel numbers them,
    in their order; a book whose parts no macro changed keeps its file's."""
    from pyopenvba._relationships import in_office_order
    from pyopenvba.powerquery._sheets import add_content_type, add_relationship, drop_content_type

    if not book.custom_xml_changed:
        return
    parts = [part for part in (book.custom_xml_parts_ or []) if not part.built_in]
    for name in [name for name in package.names() if name.startswith('customXml/')]:
        if re.fullmatch(r'customXml/itemProps\d+\.xml', name):
            drop_content_type(package, name)
        package.remove(name)
    rels = 'xl/_rels/workbook.xml.rels'
    text = package.read(rels).decode('utf-8')
    text = re.sub(r'<Relationship\b[^>]*Type="' + re.escape(ITEM_RELATIONSHIP) + r'"[^>]*/>', '', text)
    package.write(rels, text.encode('utf-8'))
    for number, part in enumerate(parts, 1):
        package.write(f'customXml/item{number}.xml',
                      ('' if part.document is None else serialized(part.document, in_file=True)).encode('utf-8'))
        package.write(f'customXml/itemProps{number}.xml',
                      properties_part(part.identity, part.schemas, part.document is None).encode('utf-8'))
        package.write(f'customXml/_rels/item{number}.xml.rels', item_relationships(number).encode('utf-8'))
        add_content_type(package, f'customXml/itemProps{number}.xml', PROPERTIES_TYPE)
        add_relationship(package, rels, ITEM_RELATIONSHIP, f'../customXml/item{number}.xml')
    package.write(rels, in_office_order(package.read(rels).decode('utf-8')).encode('utf-8'))
    book.custom_xml_changed = False
