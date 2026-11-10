"""The node tree a custom XML part holds, as Office keeps it, read and written (scripts/measure_custom_xml.py,
tests/fixtures/custom_xml*.json).

- CDATA is kept as text, joined with the text beside it; a line break is kept as a line feed.
- ``serialized`` writes the tree as CustomXMLPart.XML and CustomXMLNode.XML read it: each line feed as a carriage
  return and a line feed, an element with nothing in it closed in its start tag, &, < and > escaped in text and &,
  < and " in attribute values. An XML declaration reads as ``<?xml version="1.0"?>``; the file a save writes gives
  it its encoding, UTF-8.
- A part whose text is not well formed is error -2147467259, with the reason MSXML gives where it was measured.
"""
from __future__ import annotations

from xml.parsers import expat

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

ELEMENT, ATTRIBUTE, TEXT, INSTRUCTION, COMMENT, DOCUMENT = 1, 2, 3, 7, 8, 9
XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace'
FAILED = -2147467259


class Node:
    """One node: an element, an attribute, text, a comment, a processing instruction or the document."""

    __slots__ = ('kind', 'name', 'uri', 'value', 'parent', 'children', 'attributes', 'declarations', 'declared',
                 'view', '__weakref__')

    def __init__(self, kind: int, name: str = '', value: str = '', uri: str = '') -> None:
        self.kind, self.name, self.value, self.uri = kind, name, value, uri
        self.parent: Node | None = None
        self.children: list[Node] = []
        #: An element's attributes, namespace declarations among them in the order written; see ``declarations``.
        self.attributes: list[Node] = []
        #: Which of ``attributes`` declare a namespace, which the Attributes collection and @* leave out.
        self.declarations: set[int] = set()
        #: Whether the document began with an XML declaration.
        self.declared = False
        #: The CustomXMLNode a macro holds for it, so the same node reads as the same object.
        self.view: object = None

    @property
    def local(self) -> str:
        return self.name.rpartition(':')[2]

    @property
    def prefix(self) -> str:
        return self.name.rpartition(':')[0]

    def plain_attributes(self) -> list[Node]:
        return [one for one in self.attributes if id(one) not in self.declarations]

    def root(self) -> Node | None:
        """A document's element."""
        return next((one for one in self.children if one.kind == ELEMENT), None)

    def text(self) -> str:
        """Its text: an element's, every text node in it joined."""
        if self.kind in (ELEMENT, DOCUMENT):
            return ''.join(one.text() for one in self.children if one.kind in (ELEMENT, TEXT))
        return self.value

    def append(self, child: Node, before: Node | None = None) -> None:
        child.parent = self
        if child.kind == TEXT and before is None and self.children and self.children[-1].kind == TEXT:
            self.children[-1].value += child.value
            return
        if before is None:
            self.children.append(child)
        else:
            self.children.insert(self.children.index(before), child)

    def namespace_of(self, prefix: str) -> str:
        """The namespace a prefix, or no prefix, stands for here."""
        node: Node | None = self
        wanted = 'xmlns:' + prefix if prefix else 'xmlns'
        while node is not None and node.kind == ELEMENT:
            for one in node.attributes:
                if id(one) in node.declarations and one.name == wanted:
                    return one.value
            node = node.parent
        return XML_NAMESPACE if prefix == 'xml' else ''


def parsed(text: str) -> Node:
    """The document ``text`` holds, or error -2147467259."""
    document = Node(DOCUMENT)
    open_tags: list[Node] = [document]
    parser = expat.ParserCreate()
    parser.ordered_attributes = True

    def start(name: str, attributes: list[str]) -> None:
        element = Node(ELEMENT, name)
        parent = open_tags[-1]
        element.parent = parent
        for place in range(0, len(attributes), 2):
            key, value = attributes[place], attributes[place + 1]
            attribute = Node(ATTRIBUTE, key, value)
            attribute.parent = element
            element.attributes.append(attribute)
            if key == 'xmlns' or key.startswith('xmlns:'):
                element.declarations.add(id(attribute))
        element.uri = element.namespace_of(element.prefix)
        for attribute in element.plain_attributes():
            attribute.uri = element.namespace_of(attribute.prefix) if attribute.prefix else ''
        parent.append(element)
        open_tags.append(element)

    def end(_name: str) -> None:
        open_tags.pop()

    def characters(data: str) -> None:
        if len(open_tags) > 1:
            open_tags[-1].append(Node(TEXT, '#text', data))

    def comment(data: str) -> None:
        open_tags[-1].append(Node(COMMENT, '#comment', data))

    def instruction(target: str, data: str) -> None:
        open_tags[-1].append(Node(INSTRUCTION, target, data))

    def declaration(_version: str, _encoding: str | None, _standalone: int) -> None:
        document.declared = True

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = characters
    parser.CommentHandler = comment
    parser.ProcessingInstructionHandler = instruction
    parser.XmlDeclHandler = declaration
    try:
        parser.Parse(text, True)
    except expat.ExpatError as failure:
        raise _failure(failure, text, open_tags) from None
    return document


def _failure(failure: expat.ExpatError, text: str, open_tags: list[Node]) -> Exception:
    """MSXML's reason, where it was measured; any other is reported unsupported."""
    if failure.code == expat.errors.codes[expat.errors.XML_ERROR_TAG_MISMATCH]:
        line = text.splitlines()[failure.lineno - 1] if text.splitlines() else ''
        closing = line[failure.offset:].removeprefix('</').split('>')[0].strip()
        return error(FAILED, f"End tag '{closing}' does not match the start tag '{open_tags[-1].name}'.\n")
    if failure.code in (expat.errors.codes[expat.errors.XML_ERROR_NO_ELEMENTS],
                        expat.errors.codes[expat.errors.XML_ERROR_UNCLOSED_TOKEN]) and len(open_tags) > 1:
        return error(FAILED, f"The following tags were not closed: {', '.join(one.name for one in open_tags[1:])}.\n")
    if failure.code in (expat.errors.codes[expat.errors.XML_ERROR_SYNTAX],
                        expat.errors.codes[expat.errors.XML_ERROR_JUNK_AFTER_DOC_ELEMENT],
                        expat.errors.codes[expat.errors.XML_ERROR_INVALID_TOKEN]) and len(open_tags) == 1:
        return error(FAILED, "Invalid at the top level of the document.\n")
    return VBAUnsupportedError(f'custom XML that is not well formed ({failure}) is not implemented')


def _escaped(text: str, attribute: bool) -> str:
    text = text.replace('&', '&amp;').replace('<', '&lt;')
    return text.replace('"', '&quot;') if attribute else text.replace('>', '&gt;')


def serialized(node: Node, *, in_file: bool = False) -> str:
    """The node as XML reads it, or with ``in_file`` the document as a save writes it."""
    return _serialized(node, in_file).replace('\n', '\r\n')


def _serialized(node: Node, in_file: bool) -> str:
    if node.kind == DOCUMENT:
        lead = ''
        if node.declared:
            lead = '<?xml version="1.0" encoding="UTF-8"?>' if in_file else '<?xml version="1.0"?>'
        return lead + ''.join(_serialized(one, in_file) for one in node.children)
    if node.kind == TEXT:
        return _escaped(node.value, False)
    if node.kind == COMMENT:
        return f'<!--{node.value}-->'
    if node.kind == INSTRUCTION:
        return f'<?{node.name} {node.value}?>' if node.value else f'<?{node.name}?>'
    if node.kind == ATTRIBUTE:
        return f'{node.name}="{_escaped(node.value, True)}"'
    head = '<' + node.name + ''.join(' ' + _serialized(one, in_file) for one in node.attributes)
    if not node.children:
        return head + '/>'
    return head + '>' + ''.join(_serialized(one, in_file) for one in node.children) + f'</{node.name}>'


def document_order(node: Node) -> list[Node]:
    """Every node from ``node`` down in document order, attributes after their element."""
    found = [node]
    for attribute in node.plain_attributes():
        found.append(attribute)
    for child in node.children:
        found += document_order(child)
    return found
