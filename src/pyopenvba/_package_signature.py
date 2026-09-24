"""A VBA project's digital signature in a zip-based Office file.

Excel, Word and PowerPoint keep the signature of a .xlsm, .docm or .pptm
in parts of their own beside vbaProject.bin rather than inside it: one
part per kind, each related from the project's relationships part with a
type of its own and each given an Override in [Content_Types].xml. The
add-ins Office installs are laid out that way, and so is every file the
three applications save.

When the project's code changes, each application drops the signature on
save (scripts/measure_signature_parts.py). The parts go and so do their
Overrides; nothing else in [Content_Types].xml changes. Their
relationships go too. Excel and PowerPoint then write no relationships
part for the project, and Word keeps its part for vbaData.xml, numbered
again from rId1.
"""

from __future__ import annotations

import posixpath
import re
from collections.abc import Callable, Sequence

CONTENT_TYPES = "[Content_Types].xml"
#: The kind of signature each relationship type holds, named as SignatureInfo names them.
KINDS = {
    "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature": "legacy",
    "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile": "agile",
    "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3": "v3",
}
_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)")
_OVERRIDE = re.compile(r"<Override\b[^>]*?(?:/>|>\s*</Override>)")
_ATTRIBUTE = re.compile(r'([\w:]+)="([^"]*)"')
_ID = re.compile(r'\bId="[^"]*"')


def relationships_part(vba_entry: str) -> str:
    """The part holding the relationships of ``vba_entry``: xl/_rels/vbaProject.bin.rels for xl/vbaProject.bin."""
    folder, _, name = vba_entry.rpartition("/")
    return f"{folder}/_rels/{name}.rels" if folder else f"_rels/{name}.rels"


def signature_parts(names: Sequence[str], read: Callable[[str], bytes], vba_entry: str) -> dict[str, str]:
    """Each part the project's relationships name as holding its signature, with the kind it holds.

    ``names`` are the package's entries and ``read`` reads one. A part
    is named as the package names it; one the relationships point at
    that the package lacks keeps the name the relationship gives it.
    """
    actual = {name.casefold(): name for name in names}
    rels = actual.get(relationships_part(vba_entry).casefold())
    if rels is None:
        return {}
    folder = vba_entry.rpartition("/")[0]
    found: dict[str, str] = {}
    for element in _RELATIONSHIP.findall(read(rels).decode("utf-8")):
        attributes = dict(_ATTRIBUTE.findall(element))
        kind = KINDS.get(attributes.get("Type", ""))
        if kind is None or attributes.get("TargetMode") == "External":
            continue
        target = attributes.get("Target", "")
        part = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        found[actual.get(part.casefold(), part)] = kind
    return found


def without_signature(names: Sequence[str], read: Callable[[str], bytes], vba_entry: str) -> dict[str, bytes | None]:
    """What taking the project's signature out changes in the package, as Office takes it out.

    None marks each part that goes, and new bytes each part rewritten:
    the relationships part less the signature's relationships, numbered
    again from rId1 or gone when nothing is left in it, and
    [Content_Types].xml less the signature parts' Overrides. Empty when
    the project's relationships name no signature.
    """
    parts = signature_parts(names, read, vba_entry)
    if not parts:
        return {}
    actual = {name.casefold(): name for name in names}
    edits: dict[str, bytes | None] = {part: None for part in parts}
    rels = actual[relationships_part(vba_entry).casefold()]
    text = read(rels).decode("utf-8")
    kept = 0

    def relationship(match: re.Match[str]) -> str:
        nonlocal kept
        element = match.group()
        if KINDS.get(dict(_ATTRIBUTE.findall(element)).get("Type", "")) is not None:
            return ""
        kept += 1
        return _ID.sub(f'Id="rId{kept}"', element, count=1)

    remaining = _RELATIONSHIP.sub(relationship, text)
    edits[rels] = remaining.encode("utf-8") if kept else None
    types = actual.get(CONTENT_TYPES.casefold())
    if types is not None:
        dropped = {"/" + part.casefold() for part in parts}
        text = read(types).decode("utf-8")
        kept_types = _OVERRIDE.sub(
            lambda match: "" if dict(_ATTRIBUTE.findall(match.group())).get("PartName", "").casefold() in dropped
            else match.group(), text)
        if kept_types != text:
            edits[types] = kept_types.encode("utf-8")
    return edits
