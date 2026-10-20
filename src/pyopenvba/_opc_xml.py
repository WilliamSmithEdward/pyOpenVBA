"""Namespace-aware, byte-preserving edits to small OPC metadata parts."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import NoReturn
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


def opc_part(data: bytes, namespace: str, root_name: str) -> OpcPart:
    """Read OPC metadata without expanding declarations or changing unknown markup."""
    data.decode("utf-8")
    if b"\x00" in data:
        raise ValueError("OPC metadata must use UTF-8")
    result = OpcPart(data)
    parser = expat.ParserCreate(namespace_separator="}")
    depth = 0
    current: tuple[str, dict[str, str], int, int] | None = None

    def refuse(*_args: object) -> NoReturn:
        raise ValueError("Office XML must not contain DTDs or declared entities")

    def start(name: str, attrs: dict[str, str]) -> None:
        nonlocal depth, current
        at = parser.CurrentByteIndex
        if depth == 0:
            if name != namespace + "}" + root_name:
                raise ValueError(f"Expected OPC {root_name} namespace")
            # The parser has already validated this opening tag, including its QName.
            match = re.match(rb"<([^\s/>]+)", data[at:])
            assert match is not None
            result.qname = match.group(1).decode("utf-8")
            result.opening_end = _tag_end(data, at)
        elif depth == 1:
            current = (name, dict(attrs), at, _tag_end(data, at))
        depth += 1

    def end(_name: str) -> None:
        nonlocal depth, current
        depth -= 1
        at = parser.CurrentByteIndex
        if depth == 0:
            result.closing_start = at
        elif depth == 1 and current is not None:
            name, attrs, start_at, opening_end = current
            end_at = opening_end if data[opening_end - 2:opening_end] == b"/>" else _tag_end(data, at)
            result.children.append((name, attrs, start_at, end_at))
            current = None

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
    return result


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
