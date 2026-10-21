"""Namespace-aware, byte-preserving edits to small OPC metadata parts."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import NoReturn, overload
from xml.parsers import expat

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


@dataclass
class OpcPart:
    """Validated direct children and lexical byte positions for a narrow OPC edit."""

    data: bytes
    qname: str = ""
    opening_end: int = 0
    closing_start: int = 0
    children: list[tuple[str, dict[str, str], int, int]] = field(default_factory=lambda: [])

    def child(self, local: str, attributes: str) -> bytes:
        prefix = self.qname.rpartition(":")[0]
        return f'<{prefix + ":" if prefix else ""}{local} {attributes}/>'.encode()

    def insert(self, at: int, markup: bytes) -> bytes:
        if self.data[self.opening_end - 2:self.opening_end] == b"/>":
            end = f'</{self.qname}>'.encode()
            return self.data[:self.opening_end - 2] + b">" + markup + end + self.data[self.opening_end:]
        return self.data[:at] + markup + self.data[at:]


@overload
def opc_part(data: bytes, namespace: str, root_name: str, *, max_children: int | None = None) -> OpcPart: ...


@overload
def opc_part(data: bytes, namespace: str, root_name: str, *, container_path: tuple[str, ...], max_children: int | None = None) -> OpcPart | None: ...


def opc_part(data: bytes, namespace: str, root_name: str, *,
             container_path: tuple[str, ...] | None = None, max_children: int | None = None) -> OpcPart | None:
    """Read direct children, optionally inside a path of same-namespace containers.

    Parse the complete original document so inherited namespace bindings remain
    valid. Explicit paths return None when absent; all spans remain absolute.
    Only the bounded path and the selected container's direct children are kept.
    """
    if max_children is not None and max_children < 1:
        raise ValueError("OPC child budget must be positive")
    data.decode("utf-8")
    if b"\x00" in data:
        raise ValueError("OPC metadata must use UTF-8")
    result = OpcPart(data)
    parser = expat.ParserCreate(namespace_separator="}")
    depth = 0
    path = (root_name, *(container_path or ()))
    if any(re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.-]*", name) is None for name in path):
        raise ValueError("Invalid OPC container path")
    wanted = [namespace + "}" + name for name in path]
    target_depth = len(path) - 1
    names: list[str] = []
    found = False
    active = False
    current: tuple[str, dict[str, str], int, int] | None = None

    def refuse(*_args: object) -> NoReturn:
        raise ValueError("Office XML must not contain DTDs or declared entities")

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal depth, current, found, active
        at = parser.CurrentByteIndex
        if depth == 0 and name != wanted[0]:
            raise ValueError(f"Expected OPC {root_name} namespace")
        if depth <= target_depth:
            names.append(name)
        if depth == target_depth and names == wanted:
            if found:
                raise ValueError("Ambiguous OPC container path")
            found = active = True
            # The parser validated the original tag and its inherited namespace.
            match = re.match(rb"<([^\s/>]+)", data[at:])
            assert match is not None
            result.qname = match.group(1).decode("utf-8")
            result.opening_end = _tag_end(data, at)
        elif active and depth == target_depth + 1:
            current = (name, dict(attrs), at, _tag_end(data, at))
        depth += 1

    def end(_name: str) -> None:
        nonlocal depth, current, active
        depth -= 1
        at = parser.CurrentByteIndex
        if active and depth == target_depth:
            result.closing_start = at
            active = False
        elif active and depth == target_depth + 1 and current is not None:
            name, attrs, start_at, opening_end = current
            end_at = opening_end if data[opening_end - 2:opening_end] == b"/>" else _tag_end(data, at)
            if max_children is not None and len(result.children) >= max_children:
                raise ValueError("OPC metadata exceeds its direct-child budget")
            result.children.append((name, attrs, start_at, end_at))
            current = None
        if depth <= target_depth:
            del names[depth:]

    def declaration(_version: str, encoding: str | None, _standalone: int) -> None:
        if encoding is not None and encoding.lower().replace("-", "") != "utf8":
            raise ValueError("OPC metadata must use UTF-8")

    parser.XmlDeclHandler = declaration
    parser.StartDoctypeDeclHandler = refuse
    parser.EntityDeclHandler = refuse
    parser.ExternalEntityRefHandler = refuse
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(data, True)
    return result if found else None


def _tag_end(data: bytes, at: int) -> int:
    """End of a parser-validated tag, allowing angle brackets inside attributes."""
    quote = 0
    for index in range(at, len(data)):
        char = data[index]
        if quote:
            if char == quote:
                quote = 0
        elif char in (34, 39):
            quote = char
        elif char == 62:
            return index + 1
    raise ValueError("Unterminated XML tag")


@dataclass
class OpcElement:
    """A bounded selected subtree, with expanded names and text-presence flags."""

    name: str
    attributes: dict[str, str]
    children: list[OpcElement] = field(default_factory=lambda: [])
    has_text: bool = False
    has_opaque: bool = False

    def get(self, name: str, default: str = '') -> str:
        return self.attributes.get(name, default)


def bounded_child(data: bytes, namespace: str, root_name: str, child_name: str, *,
                  max_bytes: int, max_elements: int) -> OpcElement | None:
    """Scan the original document, retaining only one bounded direct child.

    Same-local-name nodes in other namespaces or nested containers are refused,
    rather than treated as an absent ordinary child. Unrelated sheet data is
    streamed without an ElementTree or a document-size admission limit.
    """
    if max_bytes < 1 or max_elements < 1:
        raise ValueError('Invalid metadata budget')
    if any(re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.-]*', name) is None for name in (root_name, child_name)):
        raise ValueError('Invalid OPC element name')
    if b'\x00' in data:
        raise ValueError('OPC metadata must use UTF-8')
    parser = expat.ParserCreate(encoding='utf-8', namespace_separator='}')
    depth = 0
    found: OpcElement | None = None
    stack: list[OpcElement] = []
    self_closing: list[bool] = []
    beginning = 0
    retained = 0
    expected_root, expected_child = namespace + '}' + root_name, namespace + '}' + child_name

    def refuse(*_args: object) -> NoReturn:
        raise ValueError('Office XML must not contain DTDs or declared entities')

    def budget(extra: int = 0) -> None:
        if parser.CurrentByteIndex - beginning + extra > max_bytes:
            raise ValueError('Selected OPC metadata exceeds its byte budget')

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal depth, found, beginning, retained
        level = depth
        depth += 1
        if depth > 256:
            raise ValueError('Office XML nesting exceeds its supported depth')
        if level == 0 and name != expected_root:
            raise ValueError('Unexpected OPC document root')
        if name.rsplit('}', 1)[-1] == child_name:
            if level != 1 or name != expected_child or found is not None:
                raise ValueError('Repeated, wrapped or foreign OPC metadata')
            beginning = parser.CurrentByteIndex
            found = OpcElement(name, attrs)
            stack.append(found)
            retained = 1
        elif stack:
            retained += 1
            if retained > max_elements:
                raise ValueError('Selected OPC metadata exceeds its element budget')
            node = OpcElement(name, attrs)
            stack[-1].children.append(node)
            stack.append(node)
        if stack:
            opening_end = _tag_end(data, parser.CurrentByteIndex)
            self_closing.append(data[opening_end - 2:opening_end] == b'/>')
            budget(sum(len(key) + len(value) for key, value in attrs.items()))

    def end(_name: str) -> None:
        nonlocal depth
        depth -= 1
        if stack:
            at = parser.CurrentByteIndex
            # Explicit end callbacks point at </name>; self-closing callbacks
            # already point past /> and must not charge the next sibling.
            closing_bytes = 0 if self_closing.pop() else _tag_end(data, at) - at
            budget(closing_bytes)
            stack.pop()

    def text(value: str) -> None:
        if stack:
            budget(len(value))
            if value.strip():
                stack[-1].has_text = True

    def opaque(*values: str) -> None:
        if stack:
            budget(sum(len(value) for value in values))
            stack[-1].has_opaque = True

    def declaration(_version: str, encoding: str | None, _standalone: int) -> None:
        if encoding is not None and encoding.lower().replace('-', '') != 'utf8':
            raise ValueError('OPC metadata must use UTF-8')

    parser.XmlDeclHandler = declaration
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = text
    parser.CommentHandler = opaque
    parser.ProcessingInstructionHandler = opaque
    parser.StartCdataSectionHandler = opaque
    parser.StartDoctypeDeclHandler = refuse
    parser.EntityDeclHandler = refuse
    parser.ExternalEntityRefHandler = refuse
    parser.SkippedEntityHandler = refuse
    try:
        for offset in range(0, len(data), 65_536):
            parser.Parse(data[offset:offset + 65_536], False)
            if stack and min(offset + 65_536, len(data)) - beginning > max_bytes:
                raise ValueError('Selected OPC metadata exceeds its byte budget')
        parser.Parse(b'', True)
    except expat.ExpatError as failure:
        raise ValueError('Invalid Office XML') from failure
    return found
