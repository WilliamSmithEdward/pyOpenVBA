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
def opc_part(data: bytes, namespace: str, root_name: str) -> OpcPart: ...


@overload
def opc_part(data: bytes, namespace: str, root_name: str, *, container_path: tuple[str, ...]) -> OpcPart | None: ...


def opc_part(data: bytes, namespace: str, root_name: str, *,
             container_path: tuple[str, ...] | None = None) -> OpcPart | None:
    """Read direct children, optionally inside a path of same-namespace containers.

    Parse the complete original document so inherited namespace bindings remain
    valid. Explicit paths return None when absent; all spans remain absolute.
    Only the bounded path and the selected container's direct children are kept.
    """
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
