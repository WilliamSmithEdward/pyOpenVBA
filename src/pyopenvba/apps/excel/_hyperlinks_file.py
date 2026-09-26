"""Hyperlinks in a file: the addresses Excel keeps and writes, the hyperlinks element, and its relationships.

Measured in live Excel (scripts/measure_hyperlinks.py, tests/fixtures/hyperlinks/):

- An address is tidied as it is given: spaces round it dropped, a tab in
  it dropped, split at its first # into the address and the subaddress.
  An http, https or ftp address has its scheme and host in lower case,
  its backslashes turned to slashes, a path of / where it had none, its
  . and .. segments resolved, and an escape of a letter, a digit, - . _ ~
  or a space unescaped; any other address with a scheme has only the
  scheme in lower case. file:///C:/x becomes C:\\x, and C:/x file://C:\\x.
  Anything else -- a relative path, a DOS or UNC path -- is kept as given.
- A link is written as a hyperlink element: ref, r:id for an address,
  location for a subaddress, tooltip, display and xr:uid, in that order.
  The address is its relationship's target, escaped: a space, a % that
  starts no escape, < > and " as %20 %25 %3c %3e %22. A DOS path on the
  drive the workbook is saved on goes relative to its folder; one on
  another drive, and a UNC path, go after file:///, and file://C:\\x as
  it is. Reading the target back unescapes what the tidying unescapes.
- display is the link's name. It is written for a link with no address,
  for one over several cells, and for one whose name is not the text its
  cell holds. A link read without one is named by that text.
- The relationships number the links with an address first, in the order
  the sheet lists them, then the sheet's other parts after them.
"""

from __future__ import annotations

import ntpath
import re
from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import parse_area
from pyopenvba._relationships import in_office_order
from pyopenvba._xml import attributes, escape

if TYPE_CHECKING:
    from pyopenvba.apps.excel._hyperlinks import Link
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.powerquery._opc import OpcFile

HYPERLINK_TYPE: Final = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

_SCHEME: Final = re.compile(r"([A-Za-z][A-Za-z0-9+.\-]+):")
#: The schemes whose addresses Excel tidies part by part (tests/fixtures/hyperlinks/canon.xlsx).
_HIERARCHICAL: Final = frozenset({"http", "https", "ftp"})
_DRIVE: Final = re.compile(r"[A-Za-z]:[\\/]")
_ESCAPE: Final = re.compile(r"%([0-9A-Fa-f]{2})")
#: What an escape stands for that tidying writes as itself.
_PLAIN: Final = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~ ")
_HYPERLINKS: Final = re.compile(r"<hyperlinks\b[^>]*?(?:/>|>(.*?)</hyperlinks>)", re.DOTALL)
_LINK: Final = re.compile(r"<hyperlink\b[^>]*?(?:/>|>.*?</hyperlink>)", re.DOTALL)
#: The worksheet children the hyperlinks element comes before, in the schema's order.
_AFTER: Final = re.compile(
    r"<(?:printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|"
    r"ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|"
    r"tableParts|extLst)\b|</worksheet>")
_RELATIONSHIPS_HEAD: Final = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')


# --- addresses ----------------------------------------------------------------------------------------------


def given(text: str) -> tuple[str, str, bool]:
    """An address as a macro gives it: the address as Excel keeps it, the subaddress after a #, and whether it
    had one."""
    text = text.strip(" ").replace("\t", "").replace("\r", "").replace("\n", "")
    address, hashed, sub = text.partition("#")
    return tidied(address), sub, bool(hashed)


def tidied(address: str) -> str:
    """An address as Excel keeps it once given (see the module's notes)."""
    if not address:
        return ""
    if address[:8].casefold() == "file:///" and _DRIVE.match(address, 8):
        return address[8:].replace("/", "\\")
    if _DRIVE.match(address) and address[2] == "/":
        return "file://" + address.replace("/", "\\")
    scheme = _SCHEME.match(address)
    if scheme is None:
        return address
    name = scheme.group(1).lower()
    rest = address[scheme.end():]
    if name in _HIERARCHICAL and rest.startswith("//"):
        return f"{name}://{_url_rest(rest[2:])}"
    return f"{name}:{rest}"


def _url_rest(rest: str) -> str:
    """What follows an http address's //: its host in lower case, its path resolved, both unescaped."""
    rest = rest.replace("\\", "/")
    ends = [place for place in (rest.find("/"), rest.find("?")) if place >= 0]
    end = min(ends) if ends else len(rest)
    authority, remainder = rest[:end], rest[end:]
    user, at, host = authority.rpartition("@")
    authority = f"{user}{at}{host.lower()}"
    path, question, query = remainder.partition("?")
    path = _unescaped(_resolved(path or "/"))
    return f"{authority}{path}{question}{_unescaped(query)}"


def _resolved(path: str) -> str:
    """A path with its . and .. segments taken out, as a browser takes them out."""
    kept: list[str] = []
    segments = path.split("/")
    for index, segment in enumerate(segments):
        last = index == len(segments) - 1
        if segment == ".":
            if last:
                kept.append("")
            continue
        if segment == "..":
            if len(kept) > 1:
                kept.pop()
            if last:
                kept.append("")
            continue
        kept.append(segment)
    return "/".join(kept) if kept and kept[0] == "" else "/" + "/".join(kept)


def _unescaped(text: str) -> str:
    return _ESCAPE.sub(lambda m: chr(int(m.group(1), 16)) if chr(int(m.group(1), 16)) in _PLAIN else m.group(0),
                       text)


def _file_url(address: str) -> bool:
    """Whether an address is the file://C:\\x a macro's C:/x becomes."""
    return address[:7].casefold() == "file://" and _DRIVE.match(address, 7) is not None


def saved_address(address: str, folder: str) -> str:
    """A link's address as a workbook saved in ``folder`` names it: a DOS path on the folder's drive relative to
    the folder, which a link named after its address is named by too; one on another drive as it is."""
    path = address[7:] if _file_url(address) else address
    if _DRIVE.match(path) and folder and ntpath.splitdrive(path)[0].casefold() == ntpath.splitdrive(folder)[0].casefold():
        return ntpath.relpath(path, folder).replace("\\", "/")
    return address


def follows_folder(address: str) -> bool:
    """Whether where a workbook is saved changes how it saves the address: a DOS path, which may go relative."""
    return _DRIVE.match(address[7:] if _file_url(address) else address) is not None


def target_of(address: str, folder: str) -> str:
    """The relationship target a link's address is saved as, for a workbook saved in ``folder``: a DOS path on
    another drive and a UNC path after file:///, the file://C:\\x form as it is."""
    saved = saved_address(address, folder)
    if _DRIVE.match(saved) or saved.startswith("\\\\"):
        return "file:///" + saved
    if _file_url(saved):
        return saved
    return re.sub(r'%(?![0-9A-Fa-f]{2})|[ <>"]', lambda m: "%" + format(ord(m.group(0)), "02x"), saved)


def read_address(target: str) -> str:
    """A link's address as Excel reads it from its relationship's target."""
    if target[:8].casefold() == "file:///":
        rest = target[8:]
        if rest.startswith("\\\\") or _DRIVE.match(rest):
            return rest
    if _file_url(target):
        return target[7:]
    return tidied(_unescaped(target))


# --- the hyperlinks element -----------------------------------------------------------------------------------


def read(sheet: Worksheet, xml: str, relationships: dict[str, tuple[str, str]]) -> list[Link]:
    """The sheet's links as its part lists them, each named as Excel names one it reads."""
    from pyopenvba.apps.excel._hyperlinks import Link, shown_text

    found = _HYPERLINKS.search(xml)
    if found is None or not found.group(1):
        return []
    links: list[Link] = []
    for element in _LINK.finditer(found.group(1)):
        fields = attributes(element.group(0))
        try:
            area = parse_area(fields.get("ref", ""), sheet="")
        except ValueError:
            continue
        relationship = relationships.get(fields.get("r:id", ""))
        address = read_address(relationship[1]) if relationship is not None and relationship[0] == "hyperlink" \
            else ""
        link = Link(area, address, fields.get("location", ""), fields.get("tooltip", ""), uid=fields.get("xr:uid", ""))
        if "display" in fields:
            link.display, link.given = fields["display"], True
        else:
            link.display = shown_text(sheet, link) or None
        links.append(link)
    return links


def _element(sheet: Worksheet, link: Link, relationship: str, name: str | None) -> str:
    from pyopenvba.apps.excel._hyperlinks import shown_text

    parts = [f'ref="{link.area.address(absolute=False)}"']
    if relationship:
        parts.append(f'r:id="{relationship}"')
    if link.sub:
        parts.append(f'location="{escape(link.sub)}"')
    if link.tip:
        parts.append(f'tooltip="{escape(link.tip)}"')
    several = link.area.rows * link.area.columns > 1
    if name and (not relationship or several or name != shown_text(sheet, link)):
        parts.append(f'display="{escape(name)}"')
    if link.uid:
        parts.append(f'xr:uid="{link.uid}"')
    return "<hyperlink " + " ".join(parts) + "/>"


def with_links(sheet: Worksheet, xml: str, relationships: list[str], folder: str) -> str:
    """The sheet's part with its hyperlinks element as the model has the links; ``relationships`` gives each
    link's id, "" for none, for a workbook saved in ``folder``."""
    from pyopenvba.apps.excel._hyperlinks import derived

    elements: list[str] = []
    for link, relationship in zip(sheet.hyperlinks, relationships, strict=True):
        name = link.display if link.display is not None else derived(saved_address(link.address, folder), link.sub)
        elements.append(_element(sheet, link, relationship, name))
    markup = f"<hyperlinks>{''.join(elements)}</hyperlinks>" if elements else ""
    existing = _HYPERLINKS.search(xml)
    if existing is not None:
        return xml[:existing.start()] + markup + xml[existing.end():]
    following = _AFTER.search(xml, max(xml.find("</sheetData>"), 0))
    if not markup or following is None:
        return xml
    return xml[:following.start()] + markup + xml[following.start():]


# --- the relationships ------------------------------------------------------------------------------------


def write(sheet: Worksheet, package: OpcFile, folder: str) -> None:
    """Write the sheet's links into its part and its relationships: the links' first, numbered in the order the
    sheet lists them, the sheet's other parts' after, each reference to one renumbered to match."""
    folder_part, _, name = sheet.part_name.rpartition("/")
    rels_part = f"{folder_part}/_rels/{name}.rels"
    rels = package.read(rels_part).decode("utf-8", errors="replace") if package.has(rels_part) else ""
    others: list[tuple[int, str, str]] = []
    for element in re.findall(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship>)", rels):
        fields = attributes(element)
        if fields.get("Type") == HYPERLINK_TYPE:
            continue
        number = fields.get("Id", "")
        others.append((int(number[3:]) if re.fullmatch(r"rId\d+", number) else 1 << 30, number, element))
    others.sort()
    targets = [target_of(link.address, folder) if link.address else "" for link in sheet.hyperlinks]
    ids: list[str] = []
    count = 0
    for target in targets:
        if target:
            count += 1
            ids.append(f"rId{count}")
        else:
            ids.append("")
    renamed = {old: f"rId{count + place}" for place, (_, old, _) in enumerate(others, start=1)}
    # What the model keeps of a table's or a control's relationship follows it.
    for table in sheet.tables:
        table.relationship = renamed.get(table.relationship, table.relationship)
    for shape in sheet.shapes_:
        if shape.control is not None:
            shape.control.relationship = renamed.get(shape.control.relationship, shape.control.relationship)
    xml = package.read(sheet.part_name).decode("utf-8", errors="replace")
    xml = re.sub(r'\br:id="([^"]*)"', lambda m: f'r:id="{renamed.get(m.group(1), m.group(1))}"', xml)
    package.write(sheet.part_name, with_links(sheet, xml, ids, folder).encode("utf-8"))
    elements = [f'<Relationship Id="{one}" Type="{HYPERLINK_TYPE}" Target="{escape(target)}" TargetMode="External"/>'
                for one, target in zip(ids, targets, strict=True) if one]
    elements += [re.sub(r'\bId="[^"]*"', f'Id="{renamed[old]}"', element, count=1) for _, old, element in others]
    if not elements:
        if package.has(rels_part):
            package.remove(rels_part)
        return
    package.write(rels_part, in_office_order(_RELATIONSHIPS_HEAD + "".join(elements) + "</Relationships>")
                  .encode("utf-8"))
