"""Parts that spell SpreadsheetML's namespace under a prefix, read in the spelling Excel saves them in.

A writer may spell a part as ``<x:worksheet xmlns:x="...">`` where Excel
writes ``<worksheet xmlns="...">``. The two are the same XML, and Excel
reads them alike; its save spells each part its own way, the namespace the
part's default, but for a pivot cache's records, which it keeps as the
file had them (scripts/measure_prefixed_parts.py,
tests/fixtures/prefixed_parts/). The model reads and writes its parts as
Excel spells them, so a workbook opened from such a file has those parts
spelled again first.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyopenvba.powerquery._opc import OpcFile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
#: The parts Excel saves in its own spelling, by their root element.
RESPELLED = frozenset({"workbook", "worksheet", "styleSheet", "sst", "table", "comments", "calcChain", "metadata",
                       "connections", "queryTable", "pivotTableDefinition", "pivotCacheDefinition"})
#: A comment, CDATA section, processing instruction or declaration, which name no element; or an element's tag.
_MARKUP = re.compile(
    r"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|<![^>]*>"
    r"|<(?P<close>/?)(?:(?P<prefix>[A-Za-z_][\w.-]*):)?(?P<name>[A-Za-z_][\w.-]*)"
    r"""(?P<attributes>(?:[^>"']|"[^"]*"|'[^']*')*)>""",
    re.DOTALL)
#: The markup-compatibility attributes whose values list prefixes.
_PREFIX_LISTS = re.compile(r"\b(?:Ignorable|ProcessContent|MustUnderstand|PreserveElements|PreserveAttributes|"
                           r"Requires)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')")


def unprefixed(text: str) -> str:
    """``text`` with SpreadsheetML's namespace its default, as Excel spells the part, where the root element spells
    it under a prefix the root declares and nothing else uses; any other part as it is."""
    prefix = ""
    # Each edit: where it starts and ends in ``text``, and what takes that place.
    edits: list[tuple[int, int, str]] = []
    for found in _MARKUP.finditer(text):
        name = found.group("name")
        if name is None:
            continue
        attributes = found.group("attributes") or ""
        at = found.start("attributes")
        if not prefix:
            prefix = found.group("prefix") or ""
            declared = re.search(rf"""(\s)xmlns:{re.escape(prefix)}\s*=\s*(["']){re.escape(MAIN)}\2""",
                                 attributes) if prefix else None
            if declared is None or name not in RESPELLED:
                return text
            edits.append((at + declared.start(), at + declared.end(), f'{declared.group(1)}xmlns="{MAIN}"'))
            attributes = attributes[:declared.start()] + attributes[declared.end():]
        if re.search(rf"\sxmlns:{re.escape(prefix)}\s*=", attributes):
            return text
        # A default namespace of another vocabulary, an attribute under the prefix, or the prefix named where
        # markup compatibility lists prefixes would each read otherwise once the prefix is gone.
        if re.search(r"\sxmlns\s*=", attributes) or re.search(rf"\s{re.escape(prefix)}:[\w.-]+\s*=", attributes):
            return text
        if any(prefix in (listed.group(1) or listed.group(2) or "").split()
               for listed in _PREFIX_LISTS.finditer(attributes)):
            return text
        # With no default namespace, an element named without a prefix is in none, and would come into
        # SpreadsheetML's.
        if found.group("prefix") is None:
            return text
        if found.group("prefix") == prefix:
            cut = found.start() + 1 + len(found.group("close"))
            edits.append((cut, cut + len(prefix) + 1, ""))
    if not edits:
        return text
    pieces: list[str] = []
    at = 0
    for start, end, replacement in sorted(edits):
        pieces += [text[at:start], replacement]
        at = end
    return "".join(pieces) + text[at:]


#: The prefixes the note and control readers look for in a VML part, by namespace.
_VML: dict[str, str] = {"urn:schemas-microsoft-com:vml": "v", "urn:schemas-microsoft-com:office:office": "o",
                        "urn:schemas-microsoft-com:office:excel": "x"}
_DECLARED = re.compile(r"""\s+xmlns(?::([\w.-]+))?\s*=\s*(?:"([^"]*)"|'([^']*)')""")
_ATTRIBUTE_NAME = re.compile(r"""(\s)([A-Za-z_][\w.-]*):([\w.-]+)(\s*=\s*(?:"[^"]*"|'[^']*'))""")


def vml_spelled(text: str) -> str:
    """A VML part with its elements and attributes under v:, o: and x:, each declared on the root, as Excel writes them. Excel
    reads those namespaces under any prefix; openpyxl's writer gives them a new one element by element (ns0:shape), which may
    even mean another namespace further on, so the prefixes are followed through their scopes. A part already spelled so
    is left as it is, to the byte."""
    if not any(uri in text for uri in _VML):
        return text
    scopes: list[dict[str, str]] = [{}]
    changed = False
    root_done = False

    def spelled(found: re.Match[str]) -> str:
        nonlocal changed, root_done
        name = found.group("name")
        if name is None:
            return found.group(0)
        close, prefix, attributes = found.group("close"), found.group("prefix") or "", found.group("attributes") or ""
        if close:
            bound = scopes.pop() if len(scopes) > 1 else scopes[0]
            wanted = _VML.get(bound.get(prefix, ""))
            changed = changed or (wanted is not None and wanted != prefix)
            return f"</{wanted}:{name}>" if wanted else found.group(0)
        scope = dict(scopes[-1])
        kept: list[str] = []
        at = 0
        for declared in _DECLARED.finditer(attributes):
            uri = declared.group(2) if declared.group(2) is not None else declared.group(3) or ""
            scope[declared.group(1) or ""] = uri
            kept.append(attributes[at:declared.start()])
            at = declared.end()
            if uri not in _VML:
                kept.append(declared.group(0))
            elif declared.group(1) != _VML[uri]:
                changed = True
        kept.append(attributes[at:])
        rest = "".join(kept)

        def attribute(named: re.Match[str]) -> str:
            nonlocal changed
            wanted_name = _VML.get(scope.get(named.group(2), ""))
            changed = changed or (wanted_name is not None and wanted_name != named.group(2))
            return f"{named.group(1)}{wanted_name or named.group(2)}:{named.group(3)}{named.group(4)}"

        rest = _ATTRIBUTE_NAME.sub(attribute, rest)
        wanted = _VML.get(scope.get(prefix, ""))
        changed = changed or (wanted is not None and wanted != prefix)
        tag = f"{wanted}:{name}" if wanted else (f"{prefix}:{name}" if prefix else name)
        if not root_done:
            root_done = True
            rest = "".join(f' xmlns:{short}="{uri}"' for uri, short in _VML.items()) + rest
        if not rest.endswith("/"):
            scopes.append(scope)
        return f"<{tag}{rest}>"

    respelled = _MARKUP.sub(spelled, text)
    return respelled if changed else text

def drawing_spelled(text: str) -> str:
    """A drawing with its own elements under xdr:, as the shape readers look for them. Excel reads the namespace under any prefix,
    or as the default as openpyxl writes it (tests/fixtures/tool_drawings.json). A part that binds it more than once, or binds the
    default again further in, is left as it is."""
    root = re.search(r"<(?:([A-Za-z_][\w.-]*):)?wsDr\b([^>]*)>", text)
    if root is None or root.group(1) == "xdr":
        return text
    prefix = root.group(1)
    declaration = f'xmlns{":" + prefix if prefix else ""}="{DRAWING}"'
    if declaration not in root.group(2) or text.count(f'"{DRAWING}"') != 1 or (not prefix and len(re.findall(r"\sxmlns\s*=", text)) != 1):
        return text

    def spelled(found: re.Match[str]) -> str:
        if found.group("name") is None or found.group("prefix") != prefix:
            return found.group(0)
        attributes = found.group("attributes").replace(declaration, f'xmlns:xdr="{DRAWING}"', 1)
        return f'<{found.group("close")}xdr:{found.group("name")}{attributes}>'

    return _MARKUP.sub(spelled, text)


def respell(package: OpcFile) -> None:
    """Spell each of the workbook's parts as Excel saves it, where a part spells SpreadsheetML's namespace under a
    prefix, or VML's in a way of its own."""
    for name in package.names():
        if name.startswith("xl/") and name.endswith(".vml"):
            vml = package.read(name)
            try:
                written = vml.decode("utf-8")
            except UnicodeDecodeError:
                continue
            drawn = vml_spelled(written)
            if drawn != written:
                package.write(name, drawn.encode("utf-8"))
            continue
        if not (name.startswith("xl/") and name.endswith(".xml")):
            continue
        data = package.read(name)
        head = data[:4096]
        root = re.search(rb"<(?![?!])[\w.-]+(:)?", head)
        if root is None or root.group(1) is None or MAIN.encode() not in head:
            continue
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        spelled = unprefixed(text)
        if spelled != text:
            package.write(name, spelled.encode("utf-8"))
    for name in package.names():
        if name.endswith(".rels"):
            data = package.read(name)
            given = rels_unprefixed(data)
            if given is not data:
                package.write(name, given)
        elif name.startswith("xl/") and name.endswith(".xml"):
            data = package.read(name)
            given = relationships_declared(prefixes_spelled(data))
            if name.startswith("xl/worksheets/"):
                given = addressed(given)
            if given is not data:
                package.write(name, given)


#: What the readers look for each of these vocabularies under, where a part binds it to a prefix of its own.
_USUAL_PREFIX: dict[bytes, bytes] = {b"http://schemas.openxmlformats.org/officeDocument/2006/relationships": b"r",
                                     b"http://schemas.openxmlformats.org/drawingml/2006/main": b"a",
                                     b"http://schemas.openxmlformats.org/drawingml/2006/chart": b"c"}
_BOUND = re.compile(rb"""\sxmlns:([\w.-]+)\s*=\s*["']([^"']*)["']""")


def prefixes_spelled(data: bytes) -> bytes:
    """A part with r:, a: and c: spelled so, where its root binds them to another prefix. Excel goes by the namespace."""
    bound = {prefix: uri for prefix, uri in _BOUND.findall(data[:4096])}
    renames = {prefix.decode(): _USUAL_PREFIX[uri].decode() for prefix, uri in bound.items()
               if uri in _USUAL_PREFIX and prefix != _USUAL_PREFIX[uri] and _USUAL_PREFIX[uri] not in bound}
    if not renames:
        return data
    text = data.decode("utf-8")

    def spelled(found: re.Match[str]) -> str:
        name = found.group("name")
        if name is None:
            return found.group(0)
        prefix = found.group("prefix")
        attributes = found.group("attributes") or ""
        for old, new in renames.items():
            attributes = re.sub(rf"(\s)xmlns:{re.escape(old)}(\s*=)", rf"\1xmlns:{new}\2", attributes)
            attributes = re.sub(rf"(\s){re.escape(old)}:([\w.-]+\s*=)", rf"\1{new}:\2", attributes)
        written = f"{renames.get(prefix, prefix)}:{name}" if prefix else name
        return f'<{found.group("close")}{written}{attributes}>'

    return _MARKUP.sub(spelled, text).encode("utf-8")


_RELS_ROOT = re.compile(rb"<([\w.-]+):Relationships\b")


def rels_unprefixed(data: bytes) -> bytes:
    """A relationships part with its namespace the default, as every reader of one takes it."""
    root = _RELS_ROOT.search(data, 0, 1024)
    if root is None:
        return data
    prefix = re.escape(root.group(1))
    return re.sub(rb"xmlns:" + prefix + rb"(\s*=)", rb"xmlns\1", re.sub(rb"(</?)" + prefix + rb":", rb"\1", data))


_MISSING_ADDRESS = re.compile(rb"""<(?:row|c)\b(?![^>]*\sr\s*=)[^>]*>""")
_ROW_OR_CELL = re.compile(rb"""<(row|c)\b((?:[^>"']|"[^"]*"|'[^']*')*)>""")
_ADDRESS = re.compile(rb"""\sr\s*=\s*["']([A-Z]+)?(\d+)["']""")
_FIRST_CELL = re.compile(rb"""\s*<c\b[^>]*?\sr\s*=\s*["'][A-Z]+(\d+)["']""")


def addressed(data: bytes) -> bytes:
    """A sheet with an r on each row and cell. A writer may leave them out: Excel takes a row as the one after the last, or its
    first cell's, and a cell as the column after the last (tests/fixtures/tool_books/spellings.json)."""
    if _MISSING_ADDRESS.search(data) is None:
        return data
    row = column = 0

    def filled(found: re.Match[bytes]) -> bytes:
        nonlocal row, column
        attributes = found.group(2)
        address = _ADDRESS.search(attributes)
        if found.group(1) == b"row":
            column = 0
            if address is not None:
                row = int(address.group(2))
                return found.group(0)
            first = _FIRST_CELL.match(data, found.end())
            row = int(first.group(1)) if first is not None else row + 1
            return b'<row r="%d"%s>' % (row, attributes)
        if address is not None:
            letters = address.group(1) or b"A"
            column = 0
            for letter in letters:
                column = column * 26 + letter - 64
            return found.group(0)
        column += 1
        name, rest = b"", column
        while rest:
            rest, digit = divmod(rest - 1, 26)
            name = bytes([65 + digit]) + name
        return b'<c r="%s%d"%s>' % (name, row, attributes)

    return _ROW_OR_CELL.sub(filled, data)


#: A root element whose children the model writes r:id on, as a sheet's hyperlinks and drawings and a workbook's sheets.
_RELATED_ROOT = re.compile(rb"<(?:workbook|worksheet|chartsheet)\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>")
RELATIONSHIPS = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def relationships_declared(data: bytes) -> bytes:
    """A part with the r: prefix declared on its root, as Excel writes it: openpyxl declares it on each element that uses it, so
    what the model writes in place of one of those, or adds, would name a prefix nothing binds and Excel would refuse the file.
    A part that has it, or uses the prefix for something else, is returned as it is."""
    root = _RELATED_ROOT.search(data, 0, min(len(data), 4096))
    if root is None or re.search(rb"\sxmlns:r\s*=", root.group(1)) or re.search(rb"\sxmlns:r\s*=\s*[\"'](?!" + re.escape(RELATIONSHIPS) + rb")", data):
        return data
    return data[:root.end(1)] + b' xmlns:r="' + RELATIONSHIPS + b'"' + data[root.end(1):]
