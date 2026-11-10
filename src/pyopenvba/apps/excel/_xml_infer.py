"""XML imported with no map: the schema Excel infers from the data, the map made of it, and the table or cells it fills at
Destination (tests/fixtures/xml_infer.json, measured by scripts/measure_xml_infer.py).

Excel writes the schema it infers in one fixed form: every element nillable, every one but the root with minOccurs="0", one that
occurs more than once under one parent with maxOccurs="unbounded", every sequence with minOccurs="0", and each leaf typed by the
narrowest of its types all its values take. A repeating element becomes a table at Destination, its attributes then its leaves
as the columns and their names as the headers; data that repeats nothing maps one cell to each leaf across one row. Shapes whose
map Excel cannot export -- a leaf beside the repeating element, two repeating elements, an attribute outside it, a leaf that
repeats -- and mixed content are not built.
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from pyopenvba._xml import escape
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.exceptions import VBAUnsupportedError

_XSD = "http://www.w3.org/2001/XMLSchema"


@dataclass
class _Shape:
    """One element name under one parent, merged over every place it occurs."""
    name: str
    #: How many times it occurs under one parent at most, which makes it repeat.
    most: int = 0
    children: dict[str, _Shape] = field(default_factory=lambda: dict[str, "_Shape"]())
    attributes: dict[str, list[str]] = field(default_factory=lambda: dict[str, list[str]]())
    values: list[str] = field(default_factory=lambda: list[str]())
    #: Whether some occurrence has child elements, and whether some has text beside them.
    complex: bool = False
    mixed: bool = False
    #: Whether its children come in one order everywhere it occurs; where they do not, Excel declares them in an xsd:all
    #: (tests/fixtures/xml_infer.json).
    ordered: bool = True


def _split(tag: str) -> tuple[str, str]:
    """A tag's namespace and local name."""
    if tag.startswith("{"):
        uri, name = tag[1:].split("}", 1)
        return uri, name
    return "", tag


def _merge(shape: _Shape, node: XmlElement, namespaces: set[str]) -> None:
    counts: dict[str, int] = {}
    elements = list(node)
    for key, value in node.attrib.items():
        if key == "xmlns" or key.startswith("xmlns:"):
            continue
        if key.startswith("{"):
            raise VBAUnsupportedError("inferring a schema for an attribute in a namespace is not implemented")
        shape.attributes.setdefault(key, []).append(value)
    if elements:
        shape.complex = True
        if (node.text or "").strip() or any((one.tail or "").strip() for one in elements):
            shape.mixed = True
    else:
        shape.values.append(node.text or "")
    for child in elements:
        uri, name = _split(child.tag)
        namespaces.add(uri)
        counts[name] = counts.get(name, 0) + 1
        _merge(shape.children.setdefault(name, _Shape(name)), child, namespaces)
    for name, count in counts.items():
        shape.children[name].most = max(shape.children[name].most, count)
    known = list(shape.children)
    places = [known.index(name) for name in dict.fromkeys(_split(child.tag)[1] for child in elements)]
    if places != sorted(places):
        shape.ordered = False


# --- types --------------------------------------------------------------------------------------------------------------

_INTEGER = re.compile(r"\s*[+-]?\d+\s*")
_DOUBLE = re.compile(r"\s*([+-]?\d+(\.\d*)?([eE][+-]?\d+)?|NaN|-?INF)\s*")
_DATE = re.compile(r"\s*(\d{4})-(\d{2})-(\d{2})\s*")
_DATE_TIME = re.compile(r"\s*(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\s*")
_TIME = re.compile(r"\s*(\d{2}):(\d{2}):(\d{2})\s*")


def _calendar(found: re.Match[str] | None) -> bool:
    if found is None:
        return False
    try:
        dt.date(int(found.group(1)), int(found.group(2)), int(found.group(3)))
    except ValueError:
        return False
    return True


def _clock(hours: str, minutes: str, seconds: str) -> bool:
    return int(hours) < 24 and int(minutes) < 60 and int(seconds) < 60


def _takes(kind: str, text: str) -> bool:
    if kind == "integer":
        return _INTEGER.fullmatch(text) is not None
    if kind == "boolean":
        return text.strip() in ("true", "false", "1", "0")
    if kind == "double":
        return _DOUBLE.fullmatch(text) is not None
    if kind == "date":
        return _calendar(_DATE.fullmatch(text))
    if kind == "dateTime":
        found = _DATE_TIME.fullmatch(text)
        return found is not None and _calendar(found) and _clock(found.group(4), found.group(5), found.group(6))
    if kind == "time":
        found = _TIME.fullmatch(text)
        return found is not None and _clock(found.group(1), found.group(2), found.group(3))
    return True


def leaf_type(values: list[str]) -> str:
    """The narrowest type all of a leaf's values take: an integer before a boolean, so 1 and 0 alone are integers, and a
    boolean before a double (tests/fixtures/xml_infer.json). An empty value takes none but string."""
    for kind in ("integer", "boolean", "double", "date", "dateTime", "time"):
        if values and all(_takes(kind, one) for one in values):
            return kind
    return "string"


# --- the schema's text --------------------------------------------------------------------------------------------------


def _element(shape: _Shape, top: bool, form: str) -> str:
    many = ' maxOccurs="unbounded"' if shape.most > 1 else ""
    head = '<xsd:element nillable="true"' if top else f'<xsd:element minOccurs="0"{many} nillable="true"'
    tail = "" if top else f' form="{form}"'
    if not shape.complex:
        if shape.attributes:
            raise VBAUnsupportedError("inferring a schema for an element with attributes and text is not implemented")
        return f'{head} type="xsd:{leaf_type(shape.values)}" name="{shape.name}"{tail}></xsd:element>'
    inner = "".join(_element(one, False, form) for one in shape.children.values())
    attributes = "".join(f'<xsd:attribute name="{name}" form="unqualified" type="xsd:{leaf_type(values)}"></xsd:attribute>'
                         for name, values in shape.attributes.items())
    if shape.ordered:
        group = f'<xsd:sequence minOccurs="0">{inner}</xsd:sequence>'
    else:
        if any(one.most > 1 for one in shape.children.values()):
            raise VBAUnsupportedError("inferring a schema for repeating elements that come in different orders is not implemented")
        group = f"<xsd:all>{inner}</xsd:all>"
    return f'{head} name="{shape.name}"{tail}><xsd:complexType>{group}{attributes}</xsd:complexType></xsd:element>'


def _admit(root: _Shape) -> None:
    """Refuse a shape whose map Excel cannot export, or one not measured."""
    repeating: list[_Shape] = []

    def walk(shape: _Shape, under_repeat: bool) -> None:
        if shape.mixed:
            raise VBAUnsupportedError("inferring a schema for XML with text beside elements is not implemented")
        if shape.complex and not shape.children:
            raise VBAUnsupportedError("inferring a schema for an empty element with attributes is not implemented")
        repeats = shape.most > 1
        if repeats:
            if not shape.complex:
                raise VBAUnsupportedError("inferring a schema for a leaf that repeats is not implemented")
            repeating.append(shape)
        for child in shape.children.values():
            walk(child, under_repeat or repeats)

    walk(root, False)
    if len(repeating) > 1:
        raise VBAUnsupportedError("inferring a schema with more than one repeating element is not implemented")
    if not repeating:
        # Nothing repeats: a cell for each leaf across one row, nested or not (tests/fixtures/xml_infer.json). Attributes are
        # not measured there.
        def attributed(shape: _Shape) -> bool:
            return bool(shape.attributes) or any(attributed(one) for one in shape.children.values())

        if attributed(root):
            raise VBAUnsupportedError("inferring a schema for attributes where nothing repeats is not implemented")
        return
    found = repeating[0]
    if any(one.complex for one in found.children.values()):
        raise VBAUnsupportedError("inferring a schema for a repeating element with elements inside elements is not implemented")

    # Every element on the way down to the repeating one holds only the next step: no leaves, no attributes beside it.
    def path_to(shape: _Shape) -> bool:
        if shape is found:
            return True
        below = [one for one in shape.children.values() if path_to(one)]
        if below and (shape.attributes or len(shape.children) > 1):
            raise VBAUnsupportedError("inferring a schema with values beside the repeating element is not implemented")
        return bool(below)

    path_to(root)


def schema_text(tree: XmlElement) -> str:
    """The schema Excel infers for ``tree``, as XmlSchema.XML reads it."""
    uri, name = _split(tree.tag)
    namespaces = {uri}
    root = _Shape(name, most=1)
    _merge(root, tree, namespaces)
    if len(namespaces) > 1:
        raise VBAUnsupportedError("inferring a schema for XML in more than one namespace is not implemented")
    _admit(root)
    if uri and root.attributes:
        raise VBAUnsupportedError("inferring a schema for attributes in a namespace is not implemented")
    named = escape(uri)
    opening = f'<xsd:schema xmlns:xsd="{_XSD}"' + (f' targetNamespace="{named}" xmlns:ns0="{named}"' if uri else "") + ">"
    return opening + _element(root, True, "qualified" if uri else "unqualified") + "</xsd:schema>"
