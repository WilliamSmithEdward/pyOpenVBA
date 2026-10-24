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


def respell(package: OpcFile) -> None:
    """Spell each of the workbook's parts as Excel saves it, where a part spells SpreadsheetML's namespace under a
    prefix."""
    for name in package.names():
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
